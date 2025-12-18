VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040111 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子送件電子檔整批匯入"
   ClientHeight    =   5745
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   8955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   8955
   Begin VB.CommandButton cmdOK 
      Caption         =   "進度維護"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   0
      Left            =   6270
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   1110
      Width           =   870
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "卷宗區"
      CausesValidation=   0   'False
      Height          =   345
      Index           =   13
      Left            =   7170
      Style           =   1  '圖片外觀
      TabIndex        =   21
      Top             =   1110
      Width           =   870
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Height          =   345
      Left            =   8070
      TabIndex        =   20
      Top             =   1110
      Width           =   870
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid2 
      Height          =   825
      Left            =   2880
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   1455
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   "檔案名稱                                                             "
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
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
      Height          =   4155
      Left            =   30
      TabIndex        =   18
      Top             =   1200
      Width           =   4275
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
         TabIndex        =   4
         Top             =   0
         Width           =   705
      End
      Begin VB.ListBox List1 
         Height          =   3840
         ItemData        =   "frm040111.frx":0000
         Left            =   60
         List            =   "frm040111.frx":0002
         TabIndex        =   19
         Top             =   240
         Width           =   4155
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   315
      Left            =   8580
      TabIndex        =   2
      Top             =   780
      Width           =   345
   End
   Begin VB.Frame Frame2 
      Caption         =   "未歸電子檔案件："
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
      Height          =   4155
      Left            =   4350
      TabIndex        =   15
      Top             =   1200
      Width           =   4575
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
         Left            =   1680
         TabIndex        =   5
         Top             =   0
         Width           =   705
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
         Height          =   3885
         Left            =   60
         TabIndex        =   8
         Top             =   240
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   6853
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowUserResizing=   1
         FormatString    =   "本所案號|案件性質|發文日|缺檔"
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
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   345
      Left            =   6510
      TabIndex        =   6
      Top             =   60
      Width           =   885
   End
   Begin VB.TextBox textDate 
      Height          =   264
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   0
      Top             =   480
      Width           =   1092
   End
   Begin VB.CommandButton cmdImport 
      Caption         =   "匯入(&T)"
      Height          =   345
      Left            =   5490
      TabIndex        =   3
      Top             =   60
      Width           =   885
   End
   Begin VB.FileListBox File1 
      Height          =   450
      Left            =   1380
      TabIndex        =   11
      Top             =   30
      Visible         =   0   'False
      Width           =   525
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   405
      Left            =   840
      TabIndex        =   10
      Top             =   30
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   714
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frm040111.frx":0004
   End
   Begin VB.TextBox txtPath1 
      Height          =   315
      Left            =   1500
      TabIndex        =   1
      Text            =   "C:\temp\電子送件"
      Top             =   780
      Width           =   7065
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   345
      Left            =   7530
      TabIndex        =   7
      Top             =   60
      Width           =   885
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   300
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   12
      Top             =   5280
      Width           =   8895
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   13
         Top             =   120
         Width           =   8820
      End
   End
   Begin VB.Label lblDesc 
      Caption         =   "（ex.P105116.inv.PDF）"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   6120
      TabIndex        =   23
      Top             =   540
      Width           =   2355
   End
   Begin VB.Label Label1 
      Caption         =   "檔名規則：本所案號.副檔名.PDF"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   3510
      TabIndex        =   16
      Top             =   540
      Width           =   2715
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   810
      TabIndex        =   14
      Top             =   540
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "電子檔存放路徑："
      Height          =   180
      Left            =   90
      TabIndex        =   9
      Top             =   840
      Width           =   1440
   End
End
Attribute VB_Name = "frm040111"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (無)
'Create By Sindy 2013/10/7
Option Explicit

Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iLine1 As Integer
Dim m_PrintRpt1 As Boolean, m_PrintRpt2 As Boolean
Dim m_DefaultPrinter As String
Dim SeekPrint As Integer
Dim intUpdStarRow As Integer, intUpdEndRow As Integer
Dim strUpdCP01 As String, strUpdCP02 As String, strUpdCP03 As String, strUpdCP04 As String
Dim strUpdCP09 As String, strUpdCP10 As String, strUpdCPM26 As String

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Dim dblPrevRow As Double
Public cmdState As Integer '紀錄作用按鍵
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim m_CP09 As String, m_CP10 As String, m_CPM26 As String, m_CP110 As String
Dim m_CP64 As String 'Add By Sindy 2014/1/16
'Added by Lydia 2017/08/25
Dim m_ST03T As String '使用者部門
Dim m_SK01 As String '系統別
Dim strUpdFN As String 'Added by Lydia 2017/09/05 外商預設變更名稱
Dim strErrDesc As String 'Added by Lydia 2017/09/05 區別錯誤訊息
Dim strUpdCP09List As String 'Added by Lydia 2018/12/18 記錄FCT上傳卷宗區的收文號


Private Sub Check1_Click()
   If Grd1.Rows > 1 Then
      If Grd1.TextMatrix(1, 1) = "" Then
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

'匯入
Private Sub cmdImPort_Click()
Dim fs
Dim dblFCnt As Double
Dim dblMaxWidth As Double
Dim strTotRow As String
Dim strCaseNo As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strFileName As String
Dim strErr As String, bolErr As Boolean
Dim strTemp As String
Dim strNewFN As String 'Added by Lydia 2017/09/13 FCT案號+案件性質
Dim tmpArr As Variant, intJ As Integer 'Added by  Lydia 2018/12/18

On Error GoTo ErrHand
   
   'Added by Lydia 2019/08/27 區分FCT匯入
   If m_ST03T = "F1" Then
       Call cmdImport_FCT
       Exit Sub
   End If
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   If Right(Trim(txtPath1), 1) = "\" Then txtPath1 = Left(txtPath1, Len(txtPath1) - 1)
   
   '檢查資料夾
   Set fs = CreateObject("Scripting.FileSystemObject")
   File1.path = txtPath1.Text
   File1.Refresh
   If File1.ListCount = 0 Then
      MsgBox txtPath1.Text & " 此資料夾中，尚無電子檔！"
      txtPath1.SetFocus
      Exit Sub
   End If
   Set fs = Nothing
   
   Screen.MousePointer = vbHourglass
   
   dblMaxWidth = 8820
   Text2.Width = 0
   List1.Clear
   Grid2.Clear
   Grid2.Cols = 1
   Grid2.Rows = 1
   For dblFCnt = 0 To File1.ListCount - 1
      '檔名後4碼為.PDF者才須匯入
      If UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
         '檢查檔案是否正在使用中
         If PUB_ChkFileOpening(txtPath1.Text & "\" & Trim(File1.List(dblFCnt))) = True Then
            MsgBox Trim(File1.List(dblFCnt)) & vbCrLf & "檔案正在使用中，請關閉才可執行匯入！", vbExclamation
            Screen.MousePointer = vbDefault 'Added by Lydia 2017/12/19
            Exit Sub
         End If
         
         If m_ST03T = "P1" Then 'Added by Lydia 2017/08/25 外商和外專有指定案件性質
            Grid2.AddItem Trim(File1.List(dblFCnt))
         'Added by Lydia 2017/08/25
         ElseIf m_ST03T = "F1" Then '外商
            'Modified by Lydia 2017/09/05 若不輸入案件性質一律歸101申請=> 判斷FCT案
            'If InStr(Trim(File1.List(dblFCnt)), ".101.") > 0 Or InStr(Trim(File1.List(dblFCnt)), ".102.") > 0 Or InStr(Trim(File1.List(dblFCnt)), ".501.") > 0 Then
            If UCase(Left(Trim(File1.List(dblFCnt)), 3)) = "FCT" Then
               Grid2.AddItem Trim(File1.List(dblFCnt))
            End If
         'Mark by Lydia 2018/01/16 確認外專已改變需求,不用這支匯入PDF
         'ElseIf m_ST03T = "F2" Then '外專
         '   strExc(1) = Mid(Trim(File1.List(dblFCnt)), InStr(Trim(File1.List(dblFCnt)), ".") + 1)
        '    strExc(1) = Mid(strExc(1), 1, InStr(strExc(1), ".") - 1)
        '    'FCP 排除605年費
        '    If Val(strExc(1)) > 100 And Val(strExc(1)) < 2000 And Val(strExc(1)) <> 605 Then
        '       Grid2.AddItem Trim(File1.List(dblFCnt))
        '    End If
        'end 2018/01/16
         End If
         'end 2017/08/25
      End If
   Next dblFCnt
   Grid2.col = 0
   Grid2.row = 0
   Me.Grid2.Sort = 5 '字串昇冪
   
   strTotRow = Grid2.Rows - 1
   '清空變數值
   intUpdStarRow = 0
   intUpdEndRow = 0
   strUpdCP01 = ""
   strUpdCP02 = ""
   strUpdCP03 = ""
   strUpdCP04 = ""
   strUpdCP09 = ""
   strUpdCP10 = ""
   strUpdCPM26 = ""
   strUpdFN = "" 'Added by Lydia 2017/09/05
   strUpdCP09List = "" 'Added by Lydia 2018/12/18
   For dblFCnt = 1 To strTotRow
      Text2.Width = dblMaxWidth / Val(strTotRow) * dblFCnt: DoEvents
      strErr = "": strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
      strFileName = UCase(Grid2.TextMatrix(dblFCnt, 0))
      
      '取得案號
      If InStr(strFileName, ".") > 0 Then
         strCaseNo = Trim(Left(strFileName, InStr(strFileName, ".") - 1))
         'Added by Lydia 2017/08/25
         'Modified by Lydia 2018/01/16 確認外專已改變需求,不用這支匯入PDF
         'If m_ST03T = "F1" Or m_ST03T = "F2" Then
         If m_ST03T = "F1" Then
            'Modifed by Lydia 2018/08/06 檔案命名統一為案號.案件性質.PDF (ex.7/31 因為共用卷宗匯入區已有檔案,阿蓮改用FCT27342.201.501.pdf上傳補正201)
            'strFileName = Mid(strFileName, 1, InStr(Mid(strFileName, Len(strCaseNo) + 2), ".") + Len(strCaseNo))
            If Right(strFileName, 4) <> ".PDF" Then
                strErr = convForm(CheckStr(strFileName), 30) & "，非PDF檔"
                bolErr = True
                GoTo RunSave
            Else
                'Modified by Lydia 2018/12/18 改成多筆PDF上傳 (檔名:案號.案件性質.*pdf)
                'strExc(1) = Mid(strFileName, 1, Len(strFileName) - 4)
                'If InStr(strExc(1), ".") <> InStrRev(strExc(1), ".") Then
                '    strErr = convForm(CheckStr(strFileName), 30) & "，不符檔案命名原則：案號.案件性質.PDF"
                '    bolErr = True
                '    GoTo RunSave
                'Else
                '    strFileName = strExc(1)
                'End If
                strFileName = Mid(strFileName, 1, Len(strFileName) - 4) '去掉.PDF
            End If
            'end 2018/08/06
         End If
         'end 2017/08/25
      End If
      If Left(strCaseNo, 1) = "P" Then
         strCP01 = "P"
      'Added by Lydia 2017/08/25 抓案號-系統別
      ElseIf InStr("FCT,FCP", Left(strCaseNo, 3)) > 0 Then
         strCP01 = Left(strCaseNo, 3)
      'end 2017/08/25
      Else
         strErr = convForm(CheckStr(strFileName), 30) & "系統別有誤"
         bolErr = True
         GoTo RunSave
      End If
      If InStr(strCaseNo, "-") = 0 Then
         strCP02 = Format(Mid(strCaseNo, Len(strCP01) + 1), "000000")
         strCP03 = "0"
         strCP04 = "00"
      Else
         'Modified by Lydia 2019/06/04 參考FCT-6586-T的發文無法匯入的問題
         'strCP02 = Format(Mid(strCaseNo, Len(strCP01) + 1, InStr(strCaseNo, "-") - 1 - Len(strCP01)), "000000")
         'strCP03 = Mid(strCaseNo, InStr(strCaseNo, "-") + 1, 1)
         'If InStr(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), "-") > 0 Then
         '   strCP04 = Format(Mid(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), InStr(Mid(strCaseNo, InStr(strCaseNo, "-") + 1), "-") + 1), "00")
         'Else
         '   strCP04 = "00"
         'End If
         strTemp = SystemNumber(strCaseNo, 5)
         Call ChgCaseNo(strTemp, strExc)
         strCP01 = strExc(1)
         strCP02 = strExc(2)
         strCP03 = strExc(3)
         strCP04 = strExc(4)
         'end 2019/06/04
      End If
      '檢查strCP02的長度是否為6碼且為數字
      If Len(strCP02) <> 6 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP02長度非6碼有誤"
         bolErr = True
         GoTo RunSave
      ElseIf IsNumeric(strCP02) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP02非數字型態有誤"
         bolErr = True
         GoTo RunSave
      End If
      '檢查strCP03的長度是否為1碼且為數字
      If Len(strCP03) <> 1 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP03長度非1碼有誤"
         bolErr = True
         GoTo RunSave
      'Modify By Sindy 2018/7/23 ex:FCT-040189-T-00,第3個欄位會有非數字型態資料
'      ElseIf IsNumeric(strCP03) = False Then
'         strErr = convForm(CheckStr(strFileName), 30) & "案號CP03非數字型態有誤"
'         bolErr = True
'         GoTo RunSave
      End If
      '檢查strCP04的長度是否為2碼且為數字
      If Len(strCP04) <> 2 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP04長度非2碼有誤"
         bolErr = True
         GoTo RunSave
      ElseIf IsNumeric(strCP04) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP04非數字型態有誤"
         bolErr = True
         GoTo RunSave
      End If
      
      'Added by Lydia 2017/09/05 FCT案若不輸入案件性質一律歸101申請 (外商預設變更名稱)
      'Move by Lydia 2017/09/13 strUpdFN =>strNewFN
      If m_ST03T = "F1" And strCP01 = "FCT" Then
         'Added by Lydia 2017/09/13 記錄已檢查過的檔名
         If strNewFN <> "" And ((strUpdCP01 <> "" And strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 <> strCP01 & strCP02 & strCP03 & strCP04) Or bolErr = True) Then
            strUpdFN = strNewFN
         End If
         'end 2017/09/13
         
         '有輸入案件性質
         'Modifed by Lydia 2018/08/06 檔案命名統一為案號.案件性質.PDF
         'If InStr(strFileName, ".101") > 0 Or InStr(strFileName, ".102") > 0 Or InStr(strFileName, ".501") > 0 Then
         '   strNewFN = strFileName
         ''只有案號
         'ElseIf InStr(strFileName, ".") = 0 And strCaseNo = strFileName Then
         '   strNewFN = strFileName & ".101"
         'Modified by Lydia 2018/12/18 改成多筆PDF上傳 (檔名:案號.案件性質.*pdf)
         'If InStr(strFileName, ".") = 0 Then
         'Modified by Lydia 2019/06/04 參考FCT-6586-T
         'If InStr(strFileName, Val(strCP02) & ".") = 0 Then
         If InStr(strFileName, Val(strCP02) & ".") = 0 And InStr(strFileName, Val(strCP02) & "-") = 0 Then
            strErr = convForm(CheckStr(strFileName), 30) & "，不符檔案命名原則：案號.案件性質.PDF"
            bolErr = True
            GoTo RunSave
         '非其他案件性質或中文檔名 ex.申請書
         'ElseIf InStr(strFileName, ".") > 0 Then
         Else
         'end 2018/08/06
            'Modified by Lydia 2018/06/21 檔名全部更改為自行輸入案件性質編號
            'strExc(1) = Mid(strFileName, InStr(strFileName, ".") + 1)
            ''其他案件性質要排除
           ' If InStr(strExc(1), ".") = 0 And Val(strExc(1)) = 0 Then
            '   strNewFN = strCP01 & strCP02 & IIf(strCP03 & strCP04 <> "000", strCP03 & strCP04, "") & ".101"
            'End If
            strNewFN = PUB_GetSimpleName(strFileName)
            'end 2018/06/21
         End If
      End If
      'end 2017/09/13
      'end 2017/09/05
      
RunSave:
      If (strUpdCP01 <> "" And _
          strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 <> strCP01 & strCP02 & strCP03 & strCP04) Or _
         bolErr = True Then

         If intUpdStarRow > 0 Then
            If intUpdStarRow > 0 And intUpdEndRow = 0 Then
               intUpdEndRow = intUpdStarRow
            End If
            If strUpdCP09 = "" Then
               Call GetErrText(IIf(bolErr = True, strFileName, ""))
            Else
               Call SaveFilePDF '存檔
            End If
         End If
         '清空變數值
         intUpdStarRow = 0
         intUpdEndRow = 0
         strUpdCP09 = ""
         strUpdCP10 = ""
         strUpdCPM26 = ""
         strUpdFN = "" 'Added by Lydia 2017/09/05
         If bolErr = True Then
            List1.AddItem UCase(strErr), 0: SetListScroll List1
            bolErr = False
         Else
            '讀取下一筆資料
            'Modified by Lydia 2017/09/05 外商預設變更名稱
            'Call GetUpdCP09(strCaseNo, strFileName, strCP01, strCP02, strCP03, strCP04, dblFCnt)
            'Modified by Lydia 2107/09/13 strUpdFN => strNewFN
            Call GetUpdCP09(strCaseNo, IIf(strNewFN <> "", strNewFN, strFileName), strCP01, strCP02, strCP03, strCP04, dblFCnt)
         End If
      Else
         'Modified by Lydia 2017/09/05 外商預設變更名稱
         'Call GetUpdCP09(strCaseNo, strFileName, strCP01, strCP02, strCP03, strCP04, dblFCnt)
         'Modified by Lydia 2107/09/13 strUpdFN => strNewFN
         Call GetUpdCP09(strCaseNo, IIf(strNewFN <> "", strNewFN, strFileName), strCP01, strCP02, strCP03, strCP04, dblFCnt)
      End If
   Next dblFCnt
   
   If strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 <> strCP01 & strCP02 & strCP03 & strCP04 Then
      'Modified by Lydia 2017/09/05 外商預設變更名稱
      'Call GetUpdCP09(strCaseNo, strFileName, strCP01, strCP02, strCP03, strCP04, dblFCnt - 1)
      'Modified by Lydia 2107/09/13 strUpdFN => strNewFN
      Call GetUpdCP09(strCaseNo, IIf(strNewFN <> "", strNewFN, strFileName), strCP01, strCP02, strCP03, strCP04, dblFCnt - 1)
   End If
   If intUpdStarRow > 0 Then
      If intUpdStarRow > 0 And intUpdEndRow = 0 Then
         intUpdEndRow = intUpdStarRow
      End If
      If strUpdCP09 = "" Then
         Call GetErrText(IIf(bolErr = True, strFileName, ""))
      Else
         Call SaveFilePDF '存檔
      End If
   End If
   If bolErr = True Then
      List1.AddItem UCase(strErr), 0: SetListScroll List1
      bolErr = False
   End If
   
   'Added by Lydia 2018/12/18 記錄FCT上傳卷宗區的收文號,最後再更新CP121
   'Remove by Lydia 2020/07/22 FCT案開放發文後,再自行將匯入區之電子檔整批匯入
   'strExc(10) = ""
   'If strUpdCP09List <> "" Then
   '     tmpArr = Empty
   '     tmpArr = Split(strUpdCP09List, ",")
   '     strExc(10) = strUpdCP09List
   '     cnnConnection.BeginTrans
   '     For intJ = 0 To UBound(tmpArr)
   '          If Trim(tmpArr(intJ)) <> "" Then
   '               Call UpdateCP121(Left(tmpArr(intJ), 9), Mid(tmpArr(intJ), 11), "DATA") '檢查電子送件的電子檔是否全數歸檔
   '          End If
   '     Next intJ
   '     cnnConnection.CommitTrans
   'End If
   ''end 2018/12/18
   'end 2020/07/22
   
   Text2.Width = dblMaxWidth: DoEvents
   
   Screen.MousePointer = vbDefault
   
   MsgBox "匯入完畢！"
   Call cmdQuery_Click
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description
'Added by Lydia 2018/12/18
   If strExc(10) <> "" Then
        cnnConnection.RollbackTrans
   End If
End Sub

Private Sub GetUpdCP09(strCaseNo As String, strFileName As String, _
                       strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
                       dblFCnt As Double)
Dim strConSql As String
Dim strTemp As String
Dim strCP10 As String 'Added by Lydia 2017/08/25

   If intUpdStarRow = 0 Then
      strUpdCP01 = strCP01
      strUpdCP02 = strCP02
      strUpdCP03 = strCP03
      strUpdCP04 = strCP04
      
      intUpdStarRow = dblFCnt
   Else
      intUpdEndRow = dblFCnt
   End If
   strErrDesc = "" 'Added by Lydia 2017/09/05
   
   If m_ST03T = "P1" Then 'Added by Lydia 2017/08/25 內專
    '抓取此本所案號要歸檔的文號:
    If strUpdCP09 = "" Then
       strConSql = " cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
                   " and cp27 is not null and cp57 is null" & _
                   " and cp27>=" & DBDATE(textDate) & _
                   " and cp118 is not null" & _
                   " and cp120='Y'" & _
                   " and cp121 is null"
       '1.以新申請案為優先
       strSql = "select cp09,cp10,cpm26" & _
                " From caseprogress,casepropertymap" & _
                " where" & strConSql & _
                " and cp10 in(" & NewCasePtyList & ")" & _
                " and cp01=cpm01(+) and cp10=cpm02(+)"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
       If intI = 1 Then
          strUpdCP09 = RsTemp.Fields("cp09")
          strUpdCP10 = RsTemp.Fields("cp10")
          strUpdCPM26 = "" & RsTemp.Fields("cpm26")
       End If
    End If
    If strUpdCP09 = "" Then
       '2.同一天以A類優先
       '3.不同天則檢查檔案名稱來判讀案件性質
       strSql = "select cp27,count(*)" & _
                " From caseprogress" & _
                " where" & strConSql & _
                " and cp09<'C'" & _
                " group by cp27 order by cp27 asc"
       intI = 1
       Set RsTemp = ClsLawReadRstMsg(intI, strSql)
       If intI = 1 Then
          '2.同一天以A類優先,否則才B類 and cp09<'C'
          If RsTemp.RecordCount = 1 Then
             strSql = "select cp09,cp10,cpm26" & _
                      " From caseprogress,casepropertymap" & _
                      " where" & strConSql & _
                      " and cp09<'C'" & _
                      " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                      " order by cp09 asc"
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strSql)
             If intI = 1 Then
                RsTemp.MoveFirst
                strUpdCP09 = RsTemp.Fields("cp09")
                strUpdCP10 = RsTemp.Fields("cp10")
                strUpdCPM26 = "" & RsTemp.Fields("cpm26")
             End If
          '3.不同天則檢查檔案名稱來判讀案件性質
          ElseIf RsTemp.RecordCount > 1 Then
             strSql = "select cp27,cp09,cp10,cpm03,cpm26" & _
                      " From caseprogress,casepropertymap" & _
                      " where" & strConSql & _
                      " and cp09<'C'" & _
                      " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                      " order by cp27 asc,cp09 asc"
             intI = 1
             Set RsTemp = ClsLawReadRstMsg(intI, strSql)
             If intI = 1 Then
                RsTemp.MoveFirst
                Do While Not RsTemp.EOF
                   strTemp = UCase(strCaseNo & "." & RsTemp.Fields("cpm26"))
                   If Left(strFileName, Len(strTemp)) = strTemp Then
                      strUpdCP09 = RsTemp.Fields("cp09")
                      strUpdCP10 = RsTemp.Fields("cp10")
                      strUpdCPM26 = "" & RsTemp.Fields("cpm26")
                      Exit Do
                   End If
                   RsTemp.MoveNext
                Loop
             End If
          End If
       End If
    End If
   End If 'end 2017/08/25
   
   '外商和外專
   'Modified by Lydia 2018/01/16 確認外專已改變需求,不用這支匯入PDF
   'If m_ST03T = "F1" Or m_ST03T = "F2" Then
   If m_ST03T = "F1" Then
      'Modified by Lydia 2017/09/05 去掉 and nvl(cp121,'N')='N'
      strConSql = " cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
                  " and nvl(cp57,0)=0 and cp27>=" & CompDate(2, -7, DBDATE(textDate)) & _
                  " and nvl(cp118,'N')='Y'"
      'Added by Lydia 2018/07/02 判斷傳入檔案的案件性質
      If InStr(strFileName, ".") > 0 Then
            'Added by Lydia 2018/12/18 FCT可傳入多筆PDF,所以可能會例外+副檔名(ex.案號.案件性質.POA.PDF)
            If InStr(strFileName, ".") <> InStrRev(strFileName, ".") Then
                'Modified by Lydia 2020/07/22 案號後第一組 .. ; ex. 判斷data.1.pdf
                'strExc(1) = Mid(strFileName, InStr(strFileName, ".") + 1, InStrRev(strFileName, ".") - (InStr(strFileName, ".") + 1))
                strExc(2) = InStr(strFileName, ".")
                strExc(1) = Mid(strFileName, Val(strExc(2)) + 1, InStr(Mid(strFileName, Val(strExc(2)) + 1), ".") - 1)
                'end 2020/07/22
            Else
            'end 2018/12/18
                strExc(1) = Mid(strFileName, InStrRev(strFileName, ".") + 1)
                strFileName = strFileName & ".DATA" 'Added by Lydia 2018/12/18
            End If 'end 2018/12/18
            If Val(strExc(1)) > 100 Then
                 strConSql = strConSql & " and cp10=" & CNULL(strExc(1))
                 strCP10 = strExc(1) 'Added by Lydia 2018/12/18
            End If
      End If
      'end 2018/07/02
      If m_ST03T = "F1" Then '外商
         'Modified by Lydia 2017/09/05 +CP121
         'Modified by Lydia 2018/07/02 不限案件性質,只限電子送件尚未有檔案上傳
         'strSql = "select cp01,cp09,cp10,cpm26,nvl(cp121,'N') CP121" & _
                  " From caseprogress,casepropertymap" & _
                  " where" & strConSql & _
                  " and cp10 in ('101','102','105')" & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)"
         strSql = "select cp01,cp09,cp10,cpm26,nvl(cp121,'N') CP121" & _
                  " From caseprogress,casepropertymap" & _
                  " where" & strConSql & _
                  " and cp01=cpm01(+) and cp10=cpm02(+)"
                  
      'Mark by Lydia 2018/01/16 確認外專已改變需求,不用這支匯入PDF
'      ElseIf m_ST03T = "F2" Then '外專
'         'Modified by Lydia 2017/09/05 +CP121
'         strSql = "select cp09,cp10,cpm26,nvl(cp121,'N') CP121" & _
'                  " From caseprogress,casepropertymap" & _
'                  " where" & strConSql & _
'                  " and substr(cp09,1,1)='A' and cp10 not in ('605')" & _
'                  " and cp01=cpm01(+) and cp10=cpm02(+)"
      'end 2018/01/16
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            strTemp = UCase(strCaseNo & "." & RsTemp.Fields("cp10"))
            If Left(strFileName, Len(strTemp)) = strTemp Then
               'Added by Lydia 2017/09/05 區別錯誤訊息
               'Remove by Lydia 2020/07/22 FCT案開放發文後,再自行將匯入區之電子檔整批匯入
               'If "" & RsTemp.Fields("cp121") = "Y" Then
               '   strErrDesc = "，檔案已存在"
               '   Exit Do
               'End If
               ''end 2017/09/05
               'end 2020/07/22
               strUpdCP09 = RsTemp.Fields("cp09")
               strUpdCP10 = RsTemp.Fields("cp10")
               strUpdCPM26 = "DATA" '預設副檔名:DATA
               Exit Do
            End If
            RsTemp.MoveNext
         Loop
      'Added by Lydia 2018/06/21 增加非電子送件,檔名全部更改為自行輸入案件性質編號，匯入時移入最近一道相同案件性質已有發文日者
      ElseIf InStr(strFileName, ".") > 0 Then
            'Modified by Lydia 2018/12/18
            'strExc(1) = Mid(strFileName, InStrRev(strFileName, ".") + 1)
            'If Val(strExc(1)) > 100 Then
            If Val(strCP10) > 100 Then
            'end 2018/12/18
                 'Modified by Lydia 2018/12/18 strExc(1) => strCP10
                 strSql = "select cp01,cp09,cp10,cp27,cpm26,cpp02 " & _
                             "from caseprogress,casepropertymap,casepaperpdf " & _
                             "where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
                             "and cp10= '" & strCP10 & "' and nvl(cp57,0)=0 and nvl(cp27,0)>0 and nvl(cp118,'N')='N' " & _
                             "and cp01=cpm01(+) and cp10=cpm02(+) and cp09=cpp01(+) "
                 strSql = strSql & "order by cp27 desc,cp09"
                 intI = 1
                 Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                 If intI = 1 Then
                      RsTemp.MoveFirst
                      strTemp = "" & RsTemp.Fields("cp09")
                      Do While Not RsTemp.EOF
                           If "" & RsTemp.Fields("cp09") = strTemp Then
                               'Modified by Lydia 2018/07/19 +排除已刪除的PDF
                               'If InStr(UCase("" & RsTemp.Fields("cpp02")), ".DATA.PDF") > 0 Then
                               'Modified by Lydia 2018/08/06 檔案命名統一為案號.案件性質.PDF
                               'If InStr(UCase("" & RsTemp.Fields("cpp02")), ".DATA.PDF") > 0 And InStr(UCase("" & RsTemp.Fields("cpp02")), ".PDF.DEL") = 0 Then
                               'Modified by Lydia 2018/12/18
                               'If InStr(UCase("" & RsTemp.Fields("cpp02")), "." & strExc(1) & ".DATA.PDF") > 0 And InStr(UCase("" & RsTemp.Fields("cpp02")), ".PDF.DEL") = 0 Then
                               If InStr(UCase("" & RsTemp.Fields("cpp02")), UCase(strFileName)) > 0 And InStr(UCase("" & RsTemp.Fields("cpp02")), ".PDF.DEL") = 0 Then
                                     strErrDesc = "，檔案已存在"
                                     Exit Do
                               End If
                           Else
                                Exit Do
                           End If
                           RsTemp.MoveNext
                      Loop
                      If strTemp <> "" And strErrDesc = "" Then
                           strUpdCP09 = strTemp
                           'Modified by Lydia 2018/12/18
                           'strUpdCP10 = strExc(1)
                           strUpdCP10 = strCP10
                           strUpdCPM26 = "DATA" '預設副檔名:DATA
                      End If
                 End If
            End If
      'end 2018/06/21
      End If
  
   End If

End Sub

Private Sub GetErrText(strFName As String)
Dim i As Integer
Dim strText As String
   
   '失敗時,則整卷不存
   If intUpdStarRow > 0 Then
      For i = intUpdStarRow To intUpdEndRow
         If UCase(Trim(strFName)) <> UCase(Trim(Grid2.TextMatrix(i, 0))) Then
            'Added by Lydia 2017/09/05 區別錯誤訊息
            If strErrDesc <> "" Then
               strText = convForm(CheckStr(Grid2.TextMatrix(i, 0)), 30) & strErrDesc
            Else
            'end 2017/09/05
               strText = convForm(CheckStr(Grid2.TextMatrix(i, 0)), 30) & IIf(strUpdCP09 = "", "找不到歸卷的文號，", "")
               strText = Left(strText, Len(strText) - 1)
            End If 'end 2017/09/05
            
            List1.AddItem UCase(strText), 0: SetListScroll List1
         End If
      Next
   End If
End Sub

'存檔
Private Function SaveFilePDF() As Boolean
Dim dblFCnt As Double
Dim strFileName As String
Dim strFullFileName As String
Dim stReName As String
Dim fs, f
Dim strErr As String
Dim bolSave As Boolean
Dim bolCnn As Boolean
Dim strTcp01 As String, strTcp02 As String, strTcp03 As String, strTcp04 As String
   
On Error GoTo ErrHand
   
   For dblFCnt = intUpdStarRow To intUpdEndRow
      strFileName = Grid2.TextMatrix(dblFCnt, 0)
      strFullFileName = txtPath1.Text & "\" & strFileName
      bolSave = True
      cnnConnection.BeginTrans
      bolCnn = True
      
      '檢查檔名規則
      'Modified by Lydia 2017/09/05 外商預設變更名稱
      'If PUB_ChkEmpFlowFNMRule(strUpdCP01 & "-" & strUpdCP02 & "-" & strUpdCP03 & "-" & strUpdCP04, strFileName, "Y", strUpdCP10, , , False, False, strErr) = False Then
      If PUB_ChkEmpFlowFNMRule(strUpdCP01 & "-" & strUpdCP02 & "-" & strUpdCP03 & "-" & strUpdCP04, IIf(strUpdFN <> "", strUpdFN & ".PDF", strFileName), "Y", strUpdCP10, , , False, False, strErr) = False Then
         bolSave = False
         GoTo ReadNext
      End If
      '更名
      'Modified by Lydia 2017/08/25 外商和外專的副檔名都預設為DATA
      'If PUB_GetEmpFlowReNameFile(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04, strUpdCP10, strFileName, stReName, True, 1, False, strErr) = False Then
      'Modified by Lydia 2017/09/05 外商預設變更名稱
      'If PUB_GetEmpFlowReNameFile(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04, strUpdCP10, strFileName, stReName, True, 1, False, strErr, , IIf(m_ST03T = "P1", "", "DATA")) = False Then
      'Modified by Lydia 2018/12/18 FCT改成多筆PDF上傳, 只有案件性質.PDF才加副檔名DATA
      'If PUB_GetEmpFlowReNameFile(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04, strUpdCP10, IIf(strUpdFN <> "", strUpdFN & ".PDF", strFileName), stReName, True, 1, False, strErr, , IIf(m_ST03T = "P1", "", "DATA")) = False Then
      strExc(1) = "" '預設內專P1
      If m_ST03T = "F1" Then
           If (strUpdFN <> "" And InStr(UCase(strUpdFN), "." & strUpdCP10 & ".PDF") > 0) Or (strUpdFN = "" And InStr(UCase(strFileName), "." & strUpdCP10 & ".PDF") > 0) Then
               strExc(1) = "DATA" '只有案件性質.PDF才加副檔名DATA
           End If
      End If
      If PUB_GetEmpFlowReNameFile(strUpdCP01, strUpdCP02, strUpdCP03, strUpdCP04, strUpdCP10, IIf(strUpdFN <> "", strUpdFN & ".PDF", strFileName), stReName, True, 1, False, strErr, , strExc(1)) = False Then
      'end 2018/12/18
         bolSave = False
         GoTo ReadNext
      End If
      '檢查檔案是否已存在
      strSql = "select cpp02" & _
               " From casepaperpdf" & _
               " where cpp01='" & strUpdCP09 & "'" & _
               " and upper(cpp02)='" & UCase(stReName) & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strErr = convForm(CheckStr(strFileName), 30) & "，檔案已存在"
         bolSave = False
         GoTo ReadNext
      End If
      '檢查此文號的案號是否與系統抓到的案號一致
      strSql = "select cp01,cp02,cp03,cp04" & _
               " From caseprogress" & _
               " where cp09='" & strUpdCP09 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strTcp01 = RsTemp.Fields("cp01")
         strTcp02 = RsTemp.Fields("cp02")
         strTcp03 = RsTemp.Fields("cp03")
         strTcp04 = RsTemp.Fields("cp04")
         If strTcp01 <> strUpdCP01 Or _
            strTcp02 <> strUpdCP02 Or _
            strTcp03 <> strUpdCP03 Or _
            strTcp04 <> strUpdCP04 Then
            strErr = convForm(CheckStr(strFileName), 30) & "文號" & strUpdCP09 & _
                     "本所案號" & strTcp01 & strTcp02 & strTcp03 & strTcp04 & _
                     "與系統抓到的案號" & strUpdCP01 & strUpdCP02 & strUpdCP03 & strUpdCP04 & _
                     "不一致!"
            bolSave = False
            GoTo ReadNext
         End If
      End If
      
      Set fs = CreateObject("Scripting.FileSystemObject")
      Set f = fs.GetFile(strFullFileName)
      '檔案大小為 0 KB 有誤
      If f.Size = 0 Then
         strErr = convForm(CheckStr(strFileName), 30) & MsgText(9221)
         bolSave = False
         GoTo ReadNext
      End If
      
      If SaveAttFile_PDF(strUpdCP09, strFullFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "存檔失敗！" & vbCrLf & Err.Description
         bolSave = False
         GoTo ReadNext
      End If
      
ReadNext:
      If bolSave = False Then
         cnnConnection.RollbackTrans
         bolCnn = False
         strErr = Replace(strErr, vbCrLf, "")
         List1.AddItem UCase(strErr), 0: SetListScroll List1
      Else
         'Added by Lydia 2018/12/18 記錄FCT上傳卷宗區的收文號,最後再更新CP121
         If m_ST03T = "F1" Then
              If strUpdCP09List = "" Or (strUpdCP09List <> "" And InStr(strUpdCP09List, strUpdCP09) = 0) Then
                 strUpdCP09List = strUpdCP09List & strUpdCP09 & "-" & strUpdCP10 & ","
              End If
         Else
         'end 2018/12/18
              Call UpdateCP121(strUpdCP09, strUpdCP10, strUpdCPM26) '檢查電子送件的電子檔是否全數歸檔
         End If
         'Modify By Sindy 2014/5/21 Mark 因UpdateCP121會檢查新案是否已歸足,若未歸足會重新檢核
'         'Add By Sindy 2013/11/13
'         '檢查新申請案
'         strExc(0) = "select cp09,cp10,cpm26 from caseprogress,casepropertymap" & _
'                     " where cp01='" & strUpdCP01 & "' and cp02='" & strUpdCP02 & "' and cp03='" & strUpdCP03 & "' and cp04='" & strUpdCP04 & "'" & _
'                     " and cp10 in(" & NewCasePtyList & ")" & _
'                     " and cp57 is null and cp118 is not null and cp120='Y' and cp121 is null" & _
'                     " and cp01=cpm01(+) and cp10=cpm02(+)"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            If strUpdCP09 <> RsTemp.Fields("cp09") Then
'               Call UpdateCP121(RsTemp.Fields("cp09"), RsTemp.Fields("cp10"), "" & RsTemp.Fields("cpm26"), strUpdCP09)
'            End If
'         End If
'         '2013/11/13 END
         cnnConnection.CommitTrans
         bolCnn = False
         fs.DeleteFile strFullFileName, True '刪檔
      End If
   Next dblFCnt
   
   Exit Function
   
ErrHand:
   If bolCnn = True Then
      cnnConnection.RollbackTrans
   End If
   MsgBox Err.Description
End Function

Private Sub cmdOK_Click(Index As Integer)
cmdState = Index '紀錄作用按鍵
PubShowNextData
Exit Sub
End Sub

Public Sub PubShowNextData()
Dim i As Integer, j As Integer
   
   Select Case cmdState
      Case 0 '進度維護
         Me.Enabled = False
         For i = 1 To Grd1.Rows - 1
            Grd1.col = 0
            Grd1.row = i
            If Trim(Grd1.Text) = "V" Then
               Grd1.col = 0
               Grd1.Text = ""
               For j = 0 To Grd1.Cols - 1
                    Grd1.col = j
                    Grd1.CellBackColor = QBColor(15)
               Next j
               Screen.MousePointer = vbHourglass
               frm075004_2.SetData 0, Grd1.TextMatrix(i, 8), True
               frm075004_2.SetData 1, Grd1.TextMatrix(i, 9), False
               frm075004_2.SetData 2, Grd1.TextMatrix(i, 10), False
               frm075004_2.SetData 3, Grd1.TextMatrix(i, 11), False
               frm075004_2.SetData 4, Grd1.TextMatrix(i, 5), False
               'Modify By Sindy 2018/10/9
               'frm075004_2.m_PrevFormNm = Me.Name
               frm075004_2.SetParent Me
               '2018/10/9 END
               frm075004_2.Show
               frm075004_2.QueryDB
               Me.Hide
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
            End If
         Next i
         Me.Enabled = True

      Case 13 '卷宗區
         Me.Enabled = False
         For i = 1 To Grd1.Rows - 1
            Grd1.col = 0
            Grd1.row = i
            If Trim(Grd1.Text) = "V" Then
               Grd1.col = 0
               Grd1.Text = ""
               For j = 0 To Grd1.Cols - 1
                    Grd1.col = j
                    Grd1.CellBackColor = QBColor(15)
               Next j
               Screen.MousePointer = vbHourglass
               frm100101_L.m_strKey = Grd1.TextMatrix(i, 5) '總收文號
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
         Next i
         Me.Enabled = True
   End Select
End Sub

'列印
Private Sub cmdPrint_Click()
Dim i As Integer, j As Integer
   
   '未歸電子檔案件
   If Check1.Value = 1 Then
      iLine1 = 0
      For j = 1 To Grd1.Rows - 1
         For i = 1 To 4
            strTemp(i) = ""
         Next i
         strTemp(1) = Grd1.TextMatrix(j, 1)
         strTemp(2) = Grd1.TextMatrix(j, 2)
         strTemp(3) = Grd1.TextMatrix(j, 3)
         strTemp(4) = Grd1.TextMatrix(j, 4)
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
            For i = 1 To 1
               strTemp(i) = ""
            Next i
            strTemp(1) = List1.List(j)
            If iLine1 > 52 Then
               If iLine1 > 0 Then Printer.NewPage: iLine1 = 2
            End If
            PrintDetail2 '列印明細
         Next j
      End If
      Printer.EndDoc
      Exit Sub
   End If
   
   '匯入錯誤訊息
   If Check2.Value = 1 Then
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
         PrintDetail2 '列印明細
      Next j
      Printer.EndDoc
   End If
   
End Sub

Public Sub cmdQuery_Click()
Dim rsTmp As New ADODB.Recordset
Dim ii As Integer
   
   Screen.MousePointer = vbHourglass
   
   '清空及預設欄位值
   SetGrd
   dblPrevRow = 0
   
   If m_ST03T = "P1" Then 'Added by Lydia 2017/08/25 判斷部門
        strSql = "select ' ' as V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,Decode(PA09,'000',CPM03,CPM04) as 案件性質,sqldatet(cp27) as 發文日,'' as 缺檔,CP09,CP10,CPM26,CP01,CP02,CP03,CP04,CP110,CP64" & _
                 " From caseprogress, casepropertymap, patent" & _
                 " where cp01='P'" & _
                 " and cp27 is not null and cp57 is null" & _
                 " and cp118 is not null" & _
                 " and cp120='Y'" & _
                 " and cp121 is null" & _
                 " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
                 " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                 " order by cp27 asc,cp01||cp02||cp03||cp04 asc"
            
   'Added by Lydia 2017/08/25
   ElseIf m_ST03T = "F1" Then '外商
        'Modified by Lydia 2018/07/02 不限案件性質,只限電子送件尚未有檔案上傳
        'strSql = "select ' ' as V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,Decode(tm10,'000',CPM03,CPM04) as 案件性質,sqldatet(cp27) as 發文日,'' as 缺檔,CP09,CP10,CPM26,CP01,CP02,CP03,CP04,CP110,CP64" & _
                 " From caseprogress, casepropertymap, trademark" & _
                 " where cp01 in (" & GetAddStr(m_SK01) & ")" & _
                 " and cp27 >= " & CompDate(2, -7, textDate) & " and nvl(cp57,0)=0 and nvl(cp118,'N')='Y' and nvl(cp121,'N')='N'" & _
                 " and substr(cp09,1,1)='A' and cp10 in ('101','102','105')" & _
                 " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
                 " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                 " order by cp27 asc,cp01||cp02||cp03||cp04 asc"
        strSql = "select ' ' as V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,Decode(tm10,'000',CPM03,CPM04) as 案件性質,sqldatet(cp27) as 發文日,'' as 缺檔,CP09,CP10,CPM26,CP01,CP02,CP03,CP04,CP110,CP64" & _
                 " From caseprogress, casepropertymap, trademark" & _
                 " where cp01 in (" & GetAddStr(m_SK01) & ")" & _
                 " and cp27 >= " & CompDate(2, -7, textDate) & " and nvl(cp57,0)=0 and nvl(cp118,'N')='Y' and nvl(cp121,'N')='N'" & _
                 " and substr(cp09,1,1)='A' " & _
                 " and cp01=tm01(+) and cp02=tm02(+) and cp03=tm03(+) and cp04=tm04(+)" & _
                 " and cp01=cpm01(+) and cp10=cpm02(+)" & _
                 " order by cp27 asc,cp01||cp02||cp03||cp04 asc"
   'Mark by Lydia 2018/01/16 確認外專已改變需求,不用這支匯入PDF
'   ElseIf m_ST03T = "F2" Then '外專
'        strSql = "select ' ' as V,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as 本所案號,Decode(PA09,'000',CPM03,CPM04) as 案件性質,sqldatet(cp27) as 發文日,'' as 缺檔,CP09,CP10,CPM26,CP01,CP02,CP03,CP04,CP110,CP64" & _
'                 " From caseprogress, casepropertymap, patent" & _
'                 " where cp01 in (" & GetAddStr(m_SK01) & ")" & _
'                 " and cp27 >= " & CompDate(2, -7, textDate) & " and nvl(cp57,0)=0 and nvl(cp118,'N')='Y' and nvl(cp121,'N')='N'" & _
'                 " and substr(cp09,1,1)='A' and cp10 not in ('605')" & _
'                 " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
'                 " and cp01=cpm01(+) and cp10=cpm02(+)" & _
'                 " order by cp27 asc,cp01||cp02||cp03||cp04 asc"
   'end 2018/01/16
   End If
   'end 2017/08/25
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set Grd1.Recordset = rsTmp
      
      For ii = 1 To Grd1.Rows - 1
         m_CP01 = SystemNumber(Grd1.TextMatrix(ii, 1), 1) 'Add By Sindy 2014/5/22
         m_CP02 = SystemNumber(Grd1.TextMatrix(ii, 1), 2) 'Add By Sindy 2014/5/22
         m_CP03 = SystemNumber(Grd1.TextMatrix(ii, 1), 3) 'Add By Sindy 2014/5/22
         m_CP04 = SystemNumber(Grd1.TextMatrix(ii, 1), 4) 'Add By Sindy 2014/5/22
         m_CP09 = Grd1.TextMatrix(ii, 5)
         m_CP10 = Grd1.TextMatrix(ii, 6)
         m_CPM26 = Grd1.TextMatrix(ii, 7)
         m_CP110 = Grd1.TextMatrix(ii, 12) 'Add By Sindy 2013/10/28
         m_CP64 = Grd1.TextMatrix(ii, 13) 'Add By Sindy 2014/1/16
         Grd1.TextMatrix(ii, 4) = ChkNotExistsFile '填入缺檔
         'Add By Sindy 2014/10/1 P-109044,109233,109234後續”補文件”都是以”非電子送件”方式送出,人工至卷宗區加入POA
         '因非電子送件所以卷宗區才會沒更新到CP121
         '因狀況蠻多，想說若此處檢查無缺檔狀況立即更新CP121
         'Modified by Lydia 2020/07/22 排除FCT案;  FCT案開放發文後,再自行將匯入區之電子檔整批匯入
         'If Trim(Grd1.TextMatrix(ii, 4)) = "" And m_CP09 <> "" Then 'Memo by Lydia 2018/12/18 注意FCT可多檔上傳,只要有.DATA.PDF視為無缺檔
         If Trim(Grd1.TextMatrix(ii, 4)) = "" And m_CP09 <> "" And m_CP01 <> "FCT" Then
            'Memo by Lydia 2017/09/01 外專的檔案不只一個,有待與使用者協商
            strSql = "update caseprogress set cp121='Y' where cp09='" & m_CP09 & "'"
            cnnConnection.Execute strSql
            Grd1.RowHeight(ii) = 0
         End If
         '2014/10/1 END
      Next ii
      
      '若有資料游標停在第一筆
      Grd1.Visible = False
      Grd1.col = 0
      Grd1.row = 1
      dblPrevRow = Grd1.row
      Grd1.TextMatrix(dblPrevRow, 0) = "V"
      If rsTmp.RecordCount > 0 Then
         For ii = 0 To Grd1.Cols - 1
            Grd1.col = ii
            Grd1.CellBackColor = &HFFC0C0
         Next ii
      End If
      Grd1.Visible = True
   Else
      ShowNoData
   End If
   
   Screen.MousePointer = vbDefault
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

'檢查缺檔(內專)
Public Function ChkNotExistsFile() As String
Dim strFNm As String
   
   ChkNotExistsFile = ""
   
'   'Add By Sindy 2014/1/16
'   If InStr(m_CP64, "補委任書") > 0 And m_CP10 = 補文件 Then
'      strSql = "select cpp01,cpp02" & _
'               " From casepaperpdf" & _
'               " where cpp01='" & m_CP09 & "'" & _
'               " and substr(upper(cpp02),-9)='.POA.PDF'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 0 Then '不存在
'         ChkNotExistsFile = ChkNotExistsFile & "poa,"
'      End If
'   End If
'   '2014/1/16 END
   
   '新申請案要有4個檔案(.contact/.poa/.data/.副檔名)
   '其他,有CPM26副檔名者,含.副檔名及.data要有2個
   '     無             ,至少要有一個.data檔
   strSql = "select cpp01,cpp02" & _
            " From casepaperpdf" & _
            " where cpp01='" & m_CP09 & "'" & _
            " and substr(upper(cpp02),-9)='.DATA.PDF'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then '不存在
      ChkNotExistsFile = ChkNotExistsFile & "data,"
   End If
   'Modified by Lydia 2017/08/25 + 判斷部門 and m_st03t="P1"
   'If InStr(NewCasePtyList, m_CP10) > 0 Or Trim(m_CPM26) <> "" Then
   If (InStr(NewCasePtyList, m_CP10) > 0 Or Trim(m_CPM26) <> "") And m_ST03T = "P1" Then
      
      'Add By Sindy 2014/5/22
      If m_CP10 = "307" Then '分割
         strSql = "select dc05,dc06,dc07,dc08,pa08 from divisioncase,patent" & _
                  " where dc01='" & m_CP01 & "' and dc02='" & m_CP02 & "' and dc03='" & m_CP03 & "' and dc04='" & m_CP04 & "'" & _
                  " and dc05=pa01(+) and dc06=pa02(+) and dc07=pa03(+) and dc08=pa04(+)"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            '1.發明 2.新型 3.設計
            If RsTemp.Fields("pa08") = "1" Then
               strFNm = UCase(".inv.PDF")
               m_CPM26 = "inv"
            ElseIf RsTemp.Fields("pa08") = "2" Then
               strFNm = UCase(".utl.PDF")
               m_CPM26 = "utl"
            End If
         End If
      Else
      '2014/5/22 END
         strFNm = UCase("." & Trim(m_CPM26) & ".PDF")
      End If
      strSql = "select cpp01,cpp02" & _
               " From casepaperpdf" & _
               " where cpp01='" & m_CP09 & "'" & _
               " and substr(upper(cpp02),-" & Len(strFNm) & ")='" & strFNm & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 0 Then '不存在
         'Add By Sindy 2014/5/22 新申請案可能事後程序才補進說明書
         If InStr(NewCasePtyList, m_CP10) > 0 Then
            strSql = "select cpp01,cpp02" & _
                     " From casepaperpdf,caseprogress" & _
                     " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
                     " and cpp01(+)=cp09" & _
                     " and substr(upper(cpp02),-" & Len(strFNm) & ")='" & strFNm & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then '不存在
               If InStr(UCase(ChkNotExistsFile), UCase(m_CPM26) & ",") = 0 Then 'Add By Sindy 2014/12/2
                  ChkNotExistsFile = ChkNotExistsFile & m_CPM26 & ","
               End If
            End If
         Else
            If InStr(UCase(ChkNotExistsFile), UCase(m_CPM26) & ",") = 0 Then 'Add By Sindy 2014/12/2
               ChkNotExistsFile = ChkNotExistsFile & m_CPM26 & ","
            End If
         End If
      End If
      
      If InStr(NewCasePtyList, m_CP10) > 0 Then '新申請案
         'Modify By Sindy 2013/10/28 不出名代理人則無POA檔
         If m_CP110 <> "" Then
         '2013/10/28 END
            strSql = "select cpp01,cpp02" & _
                     " From casepaperpdf" & _
                     " where cpp01='" & m_CP09 & "'" & _
                     " and substr(upper(cpp02),-8)='.POA.PDF'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then '不存在
               'Add By Sindy 2014/5/22 新申請案可能事後程序才補進POA
               strSql = "select cpp01,cpp02" & _
                        " From casepaperpdf,caseprogress" & _
                        " where cp01='" & m_CP01 & "' and cp02='" & m_CP02 & "' and cp03='" & m_CP03 & "' and cp04='" & m_CP04 & "'" & _
                        " and cpp01(+)=cp09" & _
                        " and substr(upper(cpp02),-8)='.POA.PDF'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then '不存在
                  ChkNotExistsFile = ChkNotExistsFile & "poa,"
               End If
            End If
         End If
         
         strSql = "select cpp01,cpp02" & _
                  " From casepaperpdf" & _
                  " where cpp01='" & m_CP09 & "'" & _
                  " and substr(upper(cpp02),-12)='.CONTACT.PDF'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then '不存在
            ChkNotExistsFile = ChkNotExistsFile & "contact,"
         End If
      End If
   End If
   
   If ChkNotExistsFile <> "" Then
      ChkNotExistsFile = Left(ChkNotExistsFile, Len(ChkNotExistsFile) - 1)
   End If
End Function

Private Sub Command2_Click()
Dim sFile
   
On Error GoTo ErrHnd
   
   With CommonDialog1
      .CancelError = True
      .FileName = "*.pdf"
      .Filter = "PDF檔案 (*.pdf)|*.pdf"
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
            '記錄路徑
            txtPath1.Text = sFile(0)
            'PUB_SaveLastDate Me.Name, m_ST03T, txtPath1.Text 'Added by Lydia 2017/08/25 'Mark by Lydia 2017/09/01 改回個人設定
         Else
            '記錄路徑
            If InStr(.FileName, "\") > 0 Then
                SaveSetting "TAIE", "FCP", UCase(Me.Name) & "Dir", Left(.FileName, InStrRev(.FileName, "\") - 1)
            End If
            txtPath1.Text = Left(.FileName, InStrRev(.FileName, "\") - 1)
            'PUB_SaveLastDate Me.Name, m_ST03T, txtPath1.Text 'Added by Lydia 2017/08/25 'Mark by Lydia 2017/09/01 改回個人設定
         End If
      End If
   End With
   
   Exit Sub
   
ErrHnd:
   If Err.NUMBER <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
Dim SeekPrintL As Integer
Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   
   textDate.Text = strSrvDate(2)
   
   m_DefaultPrinter = Printer.DeviceName
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      j = j + 1
      If Printer.DeviceName = m_DefaultPrinter Then
         SeekPrint = i
      End If
   Next i
   Set Printer = Printers(SeekPrint)
   
  
   If GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "") <> "" Then
      txtPath1.Text = GetSetting("TAIE", "FCP", UCase(Me.Name) & "Dir", "")
   'Added by Lydia 2017/09/01 預設個人桌面
   Else
      txtPath1.Text = PUB_Getdesktop
   'end 2017/09/01
   End If
   
   'Added by Lydia 2017/08/25 判斷操作者部門和系統別
   If Pub_StrUserSt03 <> "M51" Then
      m_ST03T = Left(Pub_StrUserSt03, 2)
   Else
      'Modified by Lydia 2018/01/16 確認外專已改變需求,不用這支匯入PDF
      'm_ST03T = UCase(InputBox("請輸入欲操作的部門代號？" & vbCrLf & "(P1:內專 F1:外商 F2:外專)"))
      m_ST03T = UCase(InputBox("請輸入欲操作的部門代號？" & vbCrLf & "(P1:內專 F1:外商)"))
      If m_ST03T = "" Then m_ST03T = "P1"
   End If
   '內專
   If m_ST03T = "P1" Then
      m_SK01 = "P"
      'txtPath1.Text = PUB_GetLastDate(Me.Name, m_ST03T) 'Mark by Lydia 2017/09/01 改回個人設定
      lblDesc.Caption = "（ex.P105116.inv.PDF）"
   '外商
   ElseIf m_ST03T = "F1" Then
      m_SK01 = "FCT"
      'txtPath1.Text = PUB_GetLastDate(Me.Name, m_ST03T) 'Mark by Lydia 2017/09/01 改回個人設定
      lblDesc.Caption = "（ex.FCT041132.101.PDF）"
   '外專
   'Mark by Lydia 2018/01/16 確認外專已改變需求,不用這支匯入PDF
   'ElseIf m_ST03T = "F2" Then
  '    m_SK01 = "FCP"
 '     'txtPath1.Text = PUB_GetLastDate(Me.Name, m_ST03T) 'Mark by Lydia 2017/09/01 改回個人設定
 '     lblDesc.Caption = "（ex.FCP051132.101.PDF）"
   'end 2018/01/16
   End If
        'Mark by Lydia 2017/09/01
        'If Mid(txtPath1, 1, 2) <> "\\" Then
        '   If Dir(txtPath1, vbDirectory) = "" Then '檢查資料夾是否存在
        '      txtPath1.Text = PUB_Getdesktop
        '   End If
        'End If
   'end 2017/08/25
   
   'Modified by Lydia 2017/08/25
   'If strUserNum = "97038" Then
   If Pub_StrUserSt03 = "M51" Then
        'txtPath1.Text = PUB_Getdesktop ''Remove by Lydia 2018/12/18
         cmdOK(0).Visible = True
   Else
         cmdOK(0).Visible = False
   End If
   Call cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040111 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow Grd1, x, y, nCol, nRow
Grd1.col = nCol
Grd1.row = nRow
End Sub

Private Sub Grd1_Click()
Dim ii As Integer

Grd1.Visible = False
If Grd1.MouseRow <> 0 And Grd1.TextMatrix(Grd1.MouseRow, 1) <> "" Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      Grd1.col = 2
      Grd1.row = dblPrevRow
      Grd1.TextMatrix(dblPrevRow, 0) = ""
      For ii = 0 To Grd1.Cols - 1
         Grd1.col = ii
         Grd1.CellBackColor = QBColor(15)
      Next ii
   End If
   '目前資料列反白
   Grd1.col = 0
   Grd1.row = Grd1.MouseRow
   dblPrevRow = Grd1.row
   Grd1.TextMatrix(Grd1.MouseRow, 0) = "V"
   For ii = 0 To Grd1.Cols - 1
      Grd1.col = ii
      Grd1.CellBackColor = &HFFC0C0
   Next ii
End If
Grd1.Visible = True
End Sub

Private Sub textDate_GotFocus()
   InverseTextBox textDate
End Sub

Private Sub textDate_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   
   Cancel = False
   If IsEmptyText(textDate) = False Then
      If CheckIsTaiwanDate(textDate, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的發文日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_GotFocus
         GoTo EXITSUB
      End If
      
      '發文日不能大於系統日
      If DBDATE(textDate) > strSrvDate(1) Then
         Cancel = True
         strMsg = "發文日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textDate_GotFocus
      End If
   End If
EXITSUB:
End Sub

Private Sub txtPath1_GotFocus()
   InverseTextBox txtPath1
End Sub

Private Function TxtValidate() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim Cancel As Boolean

TxtValidate = False

If IsEmptyText(textDate) = True Then
   strTit = "檢核資料"
   strMsg = "請輸入發文日！"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   textDate.SetFocus
   Exit Function
End If

If Me.textDate.Enabled = True Then
   Cancel = False
   textDate_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If IsEmptyText(txtPath1) = True Then
   strTit = "檢核資料"
   strMsg = "請輸入電子檔存放路徑！"
   nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   txtPath1.SetFocus
   Exit Function
End If

TxtValidate = True
End Function

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 5500
End Sub

Sub PrintTitle()
GetPleft
iLine1 = 1

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("未歸電子檔案件清單") / 2)
Printer.CurrentY = iLine1 * 300
Printer.Print "未歸電子檔案件清單"

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
Printer.Print "缺檔"

iLine1 = iLine1 + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine1 * 300
Printer.Print String(148, "-")
iLine1 = iLine1 + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
   For m_j = 1 To 4
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

Private Sub SetGrd()
   Dim arrGrd1HeadText, arrGrd1HeadWidth
   Dim iRow As Integer
   
   arrGrd1HeadText = Array("V", "本所案號", "案件性質", "發文日", "缺檔", "CP09", "CP10", "CPM26", "CP01", "CP02", "CP03", "CP04", "CP110", "CP64")
   arrGrd1HeadWidth = Array(200, 1150, 800, 800, 1300, 0, 0, 0, 0, 0, 0, 0, 0, 0)
   Grd1.Visible = False
   Grd1.Cols = UBound(arrGrd1HeadText) + 1
   For iRow = 0 To Grd1.Cols - 1
      Grd1.row = 0
      Grd1.col = iRow
      Grd1.Text = arrGrd1HeadText(iRow)
      Grd1.ColWidth(iRow) = arrGrd1HeadWidth(iRow)
      Grd1.CellAlignment = flexAlignCenterCenter
   Next
   Grd1.Visible = True
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

'Added by Lydia 2019/08/27 FCT匯入檔案 (因為FCT一道收文可收多個檔案)
Private Sub cmdImport_FCT()
Dim fs
Dim dblFCnt As Double
Dim dblMaxWidth As Double
Dim strTotRow As String
Dim strCaseNo As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strFileName As String
Dim strErr As String, bolErr As Boolean
Dim strTemp As String
Dim strNewFN As String
Dim tmpArr As Variant, intJ As Integer

On Error GoTo ErrHand
   
   '重新檢查欄位有效性
   If TxtValidate = False Then Exit Sub
   
   If Right(Trim(txtPath1), 1) = "\" Then txtPath1 = Left(txtPath1, Len(txtPath1) - 1)
   
   '檢查資料夾
   Set fs = CreateObject("Scripting.FileSystemObject")
   File1.path = txtPath1.Text
   File1.Refresh
   If File1.ListCount = 0 Then
      MsgBox txtPath1.Text & " 此資料夾中，尚無電子檔！"
      txtPath1.SetFocus
      Exit Sub
   End If
   Set fs = Nothing
   
   Screen.MousePointer = vbHourglass
   
   dblMaxWidth = 8820
   Text2.Width = 0
   List1.Clear
   Grid2.Clear
   Grid2.Cols = 1
   Grid2.Rows = 1
   For dblFCnt = 0 To File1.ListCount - 1
      '檔名後4碼為.PDF者才須匯入
      If UCase(Right(Trim(File1.List(dblFCnt)), 4)) = ".PDF" Then
         '檢查檔案是否正在使用中
         If PUB_ChkFileOpening(txtPath1.Text & "\" & Trim(File1.List(dblFCnt))) = True Then
            MsgBox Trim(File1.List(dblFCnt)) & vbCrLf & "檔案正在使用中，請關閉才可執行匯入！", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         '外商: 若不輸入案件性質一律歸101申請=> 判斷FCT案
         If UCase(Left(Trim(File1.List(dblFCnt)), 3)) = "FCT" Then
            Grid2.AddItem Trim(File1.List(dblFCnt))
         End If
      End If
   Next dblFCnt
   Grid2.col = 0
   Grid2.row = 0
   Me.Grid2.Sort = 5 '字串昇冪
   
   strTotRow = Grid2.Rows - 1
   '清空變數值
   intUpdStarRow = 0
   intUpdEndRow = 0
   strUpdCP01 = ""
   strUpdCP02 = ""
   strUpdCP03 = ""
   strUpdCP04 = ""
   strUpdCP09 = ""
   strUpdCP10 = ""
   strUpdCPM26 = ""
   strUpdCP09List = ""
   For dblFCnt = 1 To strTotRow
      Text2.Width = dblMaxWidth / Val(strTotRow) * dblFCnt: DoEvents
      strErr = "": strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
      strFileName = UCase(Grid2.TextMatrix(dblFCnt, 0))

      '取得案號
      If InStr(strFileName, ".") > 0 Then
         strCaseNo = Trim(Left(strFileName, InStr(strFileName, ".") - 1))
         '檔案命名統一為案號.案件性質.PDF (ex.7/31 因為共用卷宗匯入區已有檔案,阿蓮改用FCT27342.201.501.pdf上傳補正201)
         If Right(strFileName, 4) <> ".PDF" Then
             strErr = convForm(CheckStr(strFileName), 30) & "，非PDF檔"
             bolErr = True
             GoTo RunSave
         Else
             strFileName = Mid(strFileName, 1, Len(strFileName) - 4) '去掉.PDF
         End If
      End If
      
      If Left(strCaseNo, 1) = "P" Then
         strCP01 = "P"
      '抓案號-系統別
      ElseIf InStr("FCT,FCP", Left(strCaseNo, 3)) > 0 Then
         strCP01 = Left(strCaseNo, 3)
      Else
         strErr = convForm(CheckStr(strFileName), 30) & "系統別有誤"
         bolErr = True
         GoTo RunSave
      End If
      If InStr(strCaseNo, "-") = 0 Then
         strCP02 = Format(Mid(strCaseNo, Len(strCP01) + 1), "000000")
         strCP03 = "0"
         strCP04 = "00"
      Else
         '參考FCT-6586-T的發文無法匯入的問題
         strTemp = SystemNumber(strCaseNo, 5)
         Call ChgCaseNo(strTemp, strExc)
         strCP01 = strExc(1)
         strCP02 = strExc(2)
         strCP03 = strExc(3)
         strCP04 = strExc(4)
      End If
      '檢查strCP02的長度是否為6碼且為數字
      If Len(strCP02) <> 6 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP02長度非6碼有誤"
         bolErr = True
         GoTo RunSave
      ElseIf IsNumeric(strCP02) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP02非數字型態有誤"
         bolErr = True
         GoTo RunSave
      End If
      '檢查strCP03的長度是否為1碼且為數字
      If Len(strCP03) <> 1 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP03長度非1碼有誤"
         bolErr = True
         GoTo RunSave
      End If
      '檢查strCP04的長度是否為2碼且為數字
      If Len(strCP04) <> 2 Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP04長度非2碼有誤"
         bolErr = True
         GoTo RunSave
      ElseIf IsNumeric(strCP04) = False Then
         strErr = convForm(CheckStr(strFileName), 30) & "案號CP04非數字型態有誤"
         bolErr = True
         GoTo RunSave
      End If
   
      If InStr(strFileName, Val(strCP02) & ".") = 0 And InStr(strFileName, Val(strCP02) & "-") = 0 Then
           strErr = convForm(CheckStr(strFileName), 30) & "，不符檔案命名原則：案號.案件性質.PDF"
           bolErr = True
           GoTo RunSave
        '非其他案件性質或中文檔名 ex.申請書
      Else
           strNewFN = PUB_GetSimpleName(strFileName)
      End If
      '取得收文號
      Call GetUpdCP09(strCaseNo, IIf(strNewFN <> "", strNewFN, strFileName), strCP01, strCP02, strCP03, strCP04, dblFCnt)
      
RunSave:
      If strUpdCP09 <> "" Or bolErr = True Or strErrDesc <> "" Then
         If intUpdStarRow > 0 Then
            If intUpdStarRow > 0 And intUpdEndRow = 0 Then
               intUpdEndRow = intUpdStarRow
            End If
            If strUpdCP09 = "" Then
               Call GetErrText(strFileName)
            Else
               Call SaveFilePDF '存檔
            End If
         End If
         '清空變數值
         intUpdStarRow = 0
         intUpdEndRow = 0
         strUpdCP09 = ""
         strUpdCP10 = ""
         strUpdCPM26 = ""
         strUpdFN = ""
         If bolErr = True Then
            List1.AddItem UCase(strErr), 0: SetListScroll List1
            bolErr = False
         End If
      End If
   Next dblFCnt
   
   '記錄FCT上傳卷宗區的收文號,最後再更新CP121
   'Remove by Lydia 2020/07/22 FCT案開放發文後,再自行將匯入區之電子檔整批匯入
   'strExc(10) = ""
   'If strUpdCP09List <> "" Then
   '     tmpArr = Empty
   '     tmpArr = Split(strUpdCP09List, ",")
   '     strExc(10) = strUpdCP09List
   '     cnnConnection.BeginTrans
   '     For intJ = 0 To UBound(tmpArr)
   '          If Trim(tmpArr(intJ)) <> "" Then
   '               Call UpdateCP121(Left(tmpArr(intJ), 9), Mid(tmpArr(intJ), 11), "DATA") '檢查電子送件的電子檔是否全數歸檔
   '          End If
   '     Next intJ
   '     cnnConnection.CommitTrans
   'End If
    'end 2020/07/22
    
   Text2.Width = dblMaxWidth: DoEvents
   
   Screen.MousePointer = vbDefault
   
   MsgBox "匯入完畢！"
   Call cmdQuery_Click
   
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description

   If strExc(10) <> "" Then
        cnnConnection.RollbackTrans
   End If
End Sub

