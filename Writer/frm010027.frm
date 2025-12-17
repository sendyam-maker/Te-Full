VERSION 5.00
Begin VB.Form frm010027 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子公文維護作業"
   ClientHeight    =   5748
   ClientLeft      =   48
   ClientTop       =   372
   ClientWidth     =   8184
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5748
   ScaleWidth      =   8184
   Begin VB.Frame fraCountDown 
      Height          =   3285
      Left            =   660
      TabIndex        =   22
      Top             =   3900
      Width           =   4335
      Begin VB.CommandButton cmdCancel 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   12
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   660
         TabIndex        =   24
         Top             =   2280
         Width           =   3165
      End
      Begin VB.Timer Timer1 
         Left            =   3150
         Top             =   240
      End
      Begin VB.Label lblCount 
         AutoSize        =   -1  'True
         Caption         =   "自動執行倒數計時"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   15.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   435
         TabIndex        =   25
         Top             =   270
         Width           =   2640
      End
      Begin VB.Label lblCountDown 
         Alignment       =   2  '置中對齊
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   72
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1440
         Left            =   1125
         TabIndex        =   23
         Top             =   750
         Width           =   1350
      End
   End
   Begin VB.Timer Timer2 
      Left            =   1395
      Top             =   4350
   End
   Begin VB.CommandButton cmdAutoRun 
      Caption         =   "自動"
      Height          =   525
      Left            =   4080
      TabIndex        =   21
      Top             =   120
      Width           =   1230
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "C:\Edoc\LT"
      Top             =   1890
      Width           =   5325
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "C:\Edoc\LP"
      Top             =   1590
      Width           =   5325
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "C:\Edoc\GT"
      Top             =   1290
      Width           =   5325
   End
   Begin VB.TextBox txtOutput 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   2610
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "C:\Edoc\GP"
      Top             =   990
      Width           =   5325
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6705
      TabIndex        =   12
      Top             =   108
      Width           =   1230
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      Left            =   3960
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   2820
      Width           =   4065
   End
   Begin VB.TextBox txtExe 
      Appearance      =   0  '平面
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "C:\E-SET\IssueCMD\IssueCMD.exe"
      Top             =   690
      Width           =   6405
   End
   Begin VB.TextBox txtDate 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2340
      TabIndex        =   6
      Text            =   "1060517"
      Top             =   2835
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "列印清單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   4
      Left            =   2700
      TabIndex        =   5
      Top             =   2250
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "列印清單及附件"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   3
      Left            =   180
      TabIndex        =   4
      Top             =   2250
      Width           =   2490
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EMail"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   2
      Left            =   2700
      TabIndex        =   3
      Top             =   120
      Width           =   1230
   End
   Begin VB.CommandButton Command1 
      Caption         =   "匯入"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Index           =   1
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   1230
   End
   Begin VB.ListBox lstHistory 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1584
      ItemData        =   "frm010027.frx":0000
      Left            =   135
      List            =   "frm010027.frx":0007
      TabIndex        =   1
      Top             =   3600
      Width           =   7890
   End
   Begin VB.CommandButton Command1 
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
      Height          =   525
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   1230
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      Left            =   675
      Style           =   2  '單純下拉式
      TabIndex        =   27
      Top             =   2820
      Width           =   870
   End
   Begin VB.CheckBox Check1 
      Caption         =   "EMail含通知櫃台待簽收統計"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4050
      TabIndex        =   28
      Top             =   2250
      Value           =   1  '核取
      Width           =   3930
   End
   Begin VB.ComboBox cmbPrinter2 
      Height          =   300
      Left            =   3960
      Style           =   2  '單純下拉式
      TabIndex        =   30
      Top             =   3180
      Width           =   4065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "更新本所案號"
      Height          =   525
      Left            =   5340
      TabIndex        =   29
      Top             =   120
      Width           =   1305
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "彩色印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   2850
      TabIndex        =   31
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "系統："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   135
      TabIndex        =   26
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公文下載目的(林景郁商標)："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   180
      TabIndex        =   19
      Top             =   1950
      Width           =   2460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公文下載目的(林景郁專利)："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   17
      Top             =   1650
      Width           =   2460
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公文下載目的(閻啟泰商標)："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   1
      Left            =   180
      TabIndex        =   15
      Top             =   1356
      Width           =   2424
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公文下載目的(閻啟泰專利)："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   192
      Index           =   0
      Left            =   180
      TabIndex        =   13
      Top             =   1056
      Width           =   2424
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   10
      Top             =   2880
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公文下載程式："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   735
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "簽收日："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1620
      TabIndex        =   7
      Top             =   2880
      Width           =   780
   End
End
Attribute VB_Name = "frm010027"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/5 改成Form2.0 (無,Printer列印未改)
'Created by Morgan 2017/5/12
Option Explicit

Const cExePath As String = "C:\E-SET\IssueCMD"
Const cExeName As String = "IssueCMD.exe"

Const cEdoc As String = "C:\EDoc" '電子公文下載資料夾
'Modified by Morgan 2024/3/26
'Const cGP As String = "GP" '桂齊恆專利
'Const cGT As String = "GT" '桂齊恆商標
Const cYP As String = "YP" '閻啟泰專利
Const cYT As String = "YT" '閻啟泰商標
'end 2024/3/26
Const cLP As String = "LP" '林景郁專利
Const cLT As String = "LT" '林景郁商標



'桂齊恆智財憑證
'Removed by Morgan 2024/3/26
''Modified by Morgan 2022/1/14 檔名有更動 舊:CN=F104187291-00-桂XX.簽章憑證.pfx
'Const cPWD1 As String = "460612"
'Const cIdFile1 As String = "\\taient5\Users\digital-ID\TIPO智財憑證\桂齊恆智財憑證\F104187291-00-桂XX.簽章憑證.pfx"
''Modified by Morgan 2020/9/23 備用憑證改抓nt6
'Const cIdFile1x As String = "\\taient6\Users\digital-ID\TIPO智財憑證\桂齊恆智財憑證\F104187291-00-桂XX.簽章憑證.pfx" '備用憑證
'end 2024/3/26

'閻啟泰智財憑證
'Added by Morgan 2024/3/276
Const cPWD1 As String = "570608"
Const cIdFile1 As String = "\\taient5\Users\digital-ID\TIPO智財憑證\閻啟泰智財憑證\F129522929-00-閻XX.簽章憑證.pfx"
'Modified by Morgan 2020/9/23 備用憑證改抓nt6
Const cIdFile1x As String = "\\taient6\Users\digital-ID\TIPO智財憑證\閻啟泰智財憑證\F129522929-00-閻XX.簽章憑證.pfx" '備用憑證
'end 2024/3/26
Dim m_IdFile1 As String

'林景郁智財憑證
Const cPWD2 As String = "jerry94007"
Const cIdFile2 As String = "\\taient5\Users\digital-ID\TIPO智財憑證\林景郁智財憑證\B121581691-00-林XX.簽章憑證.pfx"
'Modified by Morgan 2020/9/23 備用憑證改抓nt6
Const cIdFile2x As String = "\\taient6\Users\digital-ID\TIPO智財憑證\林景郁智財憑證\B121581691-00-林XX.簽章憑證.pfx" '備用憑證
Dim m_IdFile2 As String

'Modified by Lydia 2024/07/22
'Const cPFeeForm As String = "\\Pat1\Fee_Form" 'P案專利申請書存放路徑
Dim cPFeeForm As String
'end 2024/07/22
Dim m_bPFeeFormFolderCheck As Boolean
Dim m_bTFeeFormFolderCheck As Boolean

Dim rsQuery As ADODB.Recordset
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim m_PdfReader As String

'列印用
Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Dim PColName() As String

Private Const ciTitleFontSize = 22
Private Const ciFontSize = 12
Private Const ciStartX = 400
Private Const ciStartY = 500
Private Const ciColGap = 200

Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim strPrinter As String
Dim m_bolNoFileCase As Boolean, m_strRptDate As String, m_strRptSys As String, m_strIUser As String, m_bolAllCase As Boolean
Dim m_bolContinue As Boolean
Dim m_bolDivList As Boolean
Dim m_bolAutoUnload As Boolean 'Added by Morgan 2017/12/5
Dim m_iPCount As Integer '附件數
Dim m_iLCount As Integer 'Added by Morgan 2019/6/18 清單數
Public m_bCalled As Boolean 'Added by Morgan 2021/7/21

Private Function fnDowload() As Boolean
   Dim stFolder As String, stTimeStamp As String, stDownTmp As String
   Dim stProgram As String, stMessage As String
   Dim stLogFileChk As String, stLogFileNew As String
   Dim bolCSV As Boolean
   'Modified by Morgan 2024/3/26
   'Dim bolGPok As Boolean, bolGTok As Boolean
   Dim bolYPok As Boolean, bolYTok As Boolean
   'end 2024/3/26
   Dim bolLPok As Boolean, bolLTok As Boolean
   Dim iLoop As Integer
   Dim iReturn As Integer
   
   
On Error GoTo ErrHnd
   
   sbAddList "刪除備份資料夾"
   stFolder = Dir(cEdoc & "\BK_*", vbDirectory)
   Do While stFolder <> ""
      KillFolder cEdoc & "\" & stFolder
      stFolder = Dir(cEdoc & "\BK_*", vbDirectory)
   Loop
   
   bolCSV = False
   stProgram = cExePath & "\" & cExeName
   sbAddList "下載程式檢查"
   
   If Dir(stProgram) <> "" Then
      sbAddList "完成", , True
   Else
      sbAddList "失敗,程式不存在", , True
      Exit Function
   End If
   
   sbAddList "憑證檢查 林景郁"
   m_IdFile2 = cIdFile2
   If fnChkFile(m_IdFile2) = True Then
      sbAddList "完成", , True
   Else
      sbAddList "失敗,憑證不存在", , True
      
      sbAddList "備用憑證檢查 林景郁"
      m_IdFile2 = cIdFile2x
      If fnChkFile(m_IdFile2) = True Then
         sbAddList "完成", , True
      Else
         sbAddList "失敗,憑證不存在", , True
         Exit Function
      End If
   End If

   'Modified by Morgan 2024/3/26 桂齊恆->閻啟泰
   sbAddList "憑證檢查 閻啟泰"
   m_IdFile1 = cIdFile1
   If fnChkFile(m_IdFile1) = True Then
      sbAddList "完成", , True
   Else
      sbAddList "失敗,憑證不存在", , True
   
      sbAddList "備用憑證檢查 閻啟泰"
      m_IdFile1 = cIdFile1x
      If fnChkFile(m_IdFile1) = True Then
         sbAddList "完成", , True
      Else
         sbAddList "失敗,憑證不存在", , True
         Exit Function
      End If
   End If
   
   '要切換目前目錄否則會執行失敗(直接開遠端的 vbp 會無法切換目錄, 但可以先開 VB 再選遠端的 vbp)
   sbAddList "切換目前目錄"
   
   ChDir cExePath
   If CurDir = cExePath Then
      sbAddList "完成", , True
   Else
      sbAddList "失敗", , True
      Exit Function
   End If
   
   'Modified by Morgan 2017/6/1 智慧局改下載中斷也會產生CSV檔
   'Modified by Morgan 2018/3/8
   iLoop = 0
   Do While iLoop < 60
      iLoop = iLoop + 1
      'Added by Morgan 2017/10/26 智慧局說如果下載失敗, 須等180秒後才能重新下載(實際上可能不止)
      If iLoop > 1 Then Sleep 300000
      
      If Not bolLPok Then
         sbAddList "下載 林景郁 專利案件(" & iLoop & ")"
         
         'Modified by Morgan 2024/9/25 改抓本機時間,因為和資料庫時間可能會有秒差導致抓不到最新的log
         'stTimeStamp = Format(ServerTime, "000000")
         stTimeStamp = Format(Now, "hhnnss")
         'end 2024/9/25
         stDownTmp = cEdoc & "\" & cLP & strSrvDate(2) & stTimeStamp
         ShellProgram cExeName, " -S P " & stDownTmp & " " & m_IdFile2 & " " & cPWD2
         Sleep 2000
         
         '有清單
         If Dir(stDownTmp & "\*.csv") <> "" Then
            sbAddList "完成", , True
            bolCSV = True
            
         '檢查log
         ElseIf Dir(cExePath & "\logs", vbDirectory) <> "" Then
            stLogFileNew = ""
            stLogFileChk = Dir(cExePath & "\logs\IssueCMDLog_" & strSrvDate(1) & "*.*")
            Do While stLogFileChk <> ""
               If Val(Right(stLogFileChk, 10)) > Val(Right(stLogFileNew, 10)) And Val(Right(stLogFileChk, 10)) >= Val(stTimeStamp) Then
                  stLogFileNew = stLogFileChk
               End If
               stLogFileChk = Dir()
            Loop
            If stLogFileNew <> "" Then
               If ChkLogFile(cExePath & "\logs\" & stLogFileNew) = True Then
                  sbAddList "完成,無可簽收案件", , True
                  bolLPok = True
                  RmDir stDownTmp
               Else
                  sbAddList "失敗,不明錯誤", , True
                  'Modified by Morgan 2018/3/16
                  '有已下載公文時才停止下載
                  'Exit Function
                  If Dir(stDownTmp & "\*_*", vbDirectory) <> "" Then
                     Exit Function
                  End If
                  'end 2018/3/16
               End If
            Else
               sbAddList "失敗,找不到最新Log檔", , True
               Exit Function
            End If
         Else
            sbAddList "失敗,Logs資料夾不存在", , True
            Exit Function
         End If
      End If
      
      If Not bolLTok Then
         sbAddList "下載 林景郁 商標案件(" & iLoop & ")"
         
         'Modified by Morgan 2024/9/25 改抓本機時間,因為和資料庫時間可能會有秒差導致抓不到最新的log
         'stTimeStamp = Format(ServerTime, "000000")
         stTimeStamp = Format(Now, "hhnnss")
         'end 2024/9/25
         
         stDownTmp = cEdoc & "\" & cLT & strSrvDate(2) & stTimeStamp
         ShellProgram cExeName, " -S T " & stDownTmp & " " & m_IdFile2 & " " & cPWD2
         Sleep 2000
         
         '有清單
         If Dir(stDownTmp & "\*.csv") <> "" Then
            sbAddList "完成", , True
            bolCSV = True
         
         '檢查log
         ElseIf Dir(cExePath & "\logs", vbDirectory) <> "" Then
            stLogFileNew = ""
            stLogFileChk = Dir(cExePath & "\logs\IssueCMDLog_" & strSrvDate(1) & "*.*")
            Do While stLogFileChk <> ""
               If Val(Right(stLogFileChk, 10)) > Val(Right(stLogFileNew, 10)) And Val(Right(stLogFileChk, 10)) >= Val(stTimeStamp) Then
                  stLogFileNew = stLogFileChk
               End If
               stLogFileChk = Dir()
            Loop
            If stLogFileNew <> "" Then
               If ChkLogFile(cExePath & "\logs\" & stLogFileNew) = True Then
                  sbAddList "完成,無可簽收案件", , True
                  bolLTok = True
                  RmDir stDownTmp
               Else
                  sbAddList "失敗,不明錯誤", , True
                  'Modified by Morgan 2018/3/16
                  '有已下載公文時才停止下載
                  'Exit Function
                  If Dir(stDownTmp & "\*_*", vbDirectory) <> "" Then
                     Exit Function
                  End If
                  'end 2018/3/16
               End If
            Else
               sbAddList "失敗,找不到最新Log檔", , True
               Exit Function
            End If
         Else
            sbAddList "失敗,Logs資料夾不存在", , True
            Exit Function
         End If
      End If
      
      'Modified by Morgan 2024/3/26 桂齊恆->閻啟泰
      If Not bolYPok Then
         sbAddList "下載 閻啟泰 專利案件(" & iLoop & ")"
         
         'Modified by Morgan 2024/9/25 改抓本機時間,因為和資料庫時間可能會有秒差導致抓不到最新的log
         'stTimeStamp = Format(ServerTime, "000000")
         stTimeStamp = Format(Now, "hhnnss")
         'end 2024/9/25
         
         stDownTmp = cEdoc & "\" & cYP & strSrvDate(2) & stTimeStamp
         ShellProgram cExeName, " -S P " & stDownTmp & " " & m_IdFile1 & " " & cPWD1
         Sleep 2000 '至少要等1秒,否則因 log 檔名相同會寫在同一個檔案內
         '有清單
         If Dir(stDownTmp & "\*.csv") <> "" Then
            sbAddList "完成", , True
            bolCSV = True
         '無清單,檢查log
         ElseIf Dir(cExePath & "\logs", vbDirectory) <> "" Then
            stLogFileNew = ""
            stLogFileChk = Dir(cExePath & "\logs\IssueCMDLog_" & strSrvDate(1) & "*.*")
            Do While stLogFileChk <> ""
               If Val(Right(stLogFileChk, 10)) > Val(Right(stLogFileNew, 10)) And Val(Right(stLogFileChk, 10)) >= Val(stTimeStamp) Then
                  stLogFileNew = stLogFileChk
               End If
               stLogFileChk = Dir()
            Loop
            If stLogFileNew <> "" Then
               If ChkLogFile(cExePath & "\logs\" & stLogFileNew, iReturn) = True Then
                  sbAddList "完成,無可簽收案件", , True
                  'Modified by Morgan 2024/3/26
                  'bolYPok = True
                  bolYPok = True
                  'end 2024/3/26
                  RmDir stDownTmp
               'Added by Morgan 2024/3/26
               ElseIf iReturn = 1 Then
                  sbAddList "失敗,未約定電子送達", , True
                  bolYPok = True
                  RmDir stDownTmp
               'end 2024/3/26
               Else
                  sbAddList "失敗,不明錯誤", , True
                  'Modified by Morgan 2017/7/10
                  '有已下載公文時才停止下載
                  'Exit Function
                  If Dir(stDownTmp & "\*_*", vbDirectory) <> "" Then
                     Exit Function
                  End If
                  'end 2017/7/10
               End If
            Else
               sbAddList "失敗,找不到最新Log檔", , True
               Exit Function
            End If
         Else
            sbAddList "失敗,Logs資料夾不存在", , True
            Exit Function
         End If
      End If
      
      'Modified by Morgan 2024/3/26 桂齊恆->閻啟泰
      If Not bolYTok Then
         sbAddList "下載 閻啟泰 商標案件(" & iLoop & ")"
         
         'Modified by Morgan 2024/9/25 改抓本機時間,因為和資料庫時間可能會有秒差導致抓不到最新的log
         'stTimeStamp = Format(ServerTime, "000000")
         stTimeStamp = Format(Now, "hhnnss")
         'end 2024/9/25
         
         stDownTmp = cEdoc & "\" & cYT & strSrvDate(2) & stTimeStamp
         ShellProgram cExeName, " -S T " & stDownTmp & " " & m_IdFile1 & " " & cPWD1
         Sleep 2000
         
         '有清單
         If Dir(stDownTmp & "\*.csv") <> "" Then
            sbAddList "完成", , True
            bolCSV = True
            
         '檢查log
         ElseIf Dir(cExePath & "\logs", vbDirectory) <> "" Then
            stLogFileNew = ""
            stLogFileChk = Dir(cExePath & "\logs\IssueCMDLog_" & strSrvDate(1) & "*.*")
            Do While stLogFileChk <> ""
               If Val(Right(stLogFileChk, 10)) > Val(Right(stLogFileNew, 10)) And Val(Right(stLogFileChk, 10)) >= Val(stTimeStamp) Then
                  stLogFileNew = stLogFileChk
               End If
               stLogFileChk = Dir()
            Loop
            If stLogFileNew <> "" Then
               If ChkLogFile(cExePath & "\logs\" & stLogFileNew, iReturn) = True Then
                  sbAddList "完成,無可簽收案件", , True
                  bolYTok = True
                  RmDir stDownTmp
               'Added by Morgan 2024/3/26
               ElseIf iReturn = 1 Then
                  sbAddList "失敗,未約定電子送達", , True
                  RmDir stDownTmp
                  bolYTok = True
               'end 2024/3/26
               Else
                  sbAddList "失敗,不明錯誤", , True
                  'Modified by Morgan 2018/3/16
                  '有已下載公文時才停止下載
                  'Exit Function
                  If Dir(stDownTmp & "\*_*", vbDirectory) <> "" Then
                     Exit Function
                  End If
                  'end 2018/3/16
               End If
            Else
               sbAddList "失敗,找不到最新Log檔", , True
               Exit Function
            End If
         Else
            sbAddList "失敗,Logs資料夾不存在", , True
            Exit Function
         End If
      End If
            
      If bolYPok And bolYTok And bolLPok And bolLTok Then Exit Do
   Loop
   
   If bolCSV = True Then
      fnDowload = True
   Else
      sbAddList "公文下載...結束(無可簽收案件)"
   End If
   
   'ShellProgram "test.bat", ""
   ChDir App.path
   Exit Function
   
ErrHnd:
   sbAddList "" & Err.Number & "," & Err.Description
End Function
Private Function ShellProgram(ByVal program_name As String, parameters As String) As Boolean
   Dim stCmd As String
   Dim process_id As Long
   Dim process_handle As Long
      
On Error GoTo ShellError

    stCmd = program_name & " " & parameters
    process_id = Shell(stCmd, vbNormalFocus)
    DoEvents
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If

    ShellProgram = True
    Exit Function

ShellError:
   sbAddList "" & Err.Number & "," & Err.Description & "(" & stCmd & ")"
End Function

Private Function fnImport() As Boolean
   Dim stFolder As String, stCSV As String
   
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
      If MsgBox("目前連線為測試資料庫，是否確定要匯入？", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
   End If
   
   '閻啟泰-專利
   stFolder = Dir(cEdoc & "\" & cYP & strSrvDate(2) & "*", vbDirectory)
   Do While stFolder <> ""
      sbAddList "匯入 閻啟泰 專利(" & stFolder & ")"
      stCSV = Dir(cEdoc & "\" & stFolder & "\*.csv")
      If stCSV <> "" Then
         If Import2DB(cEdoc & "\" & stFolder & "\" & stCSV) = True Then
            Name cEdoc & "\" & stFolder As cEdoc & "\" & "BK_" & stFolder '移到備份資料夾以便測試
         Else
            Exit Function
         End If
      'Added by Morgan 2017/11/8 若是空資料夾則刪除後繼續
      ElseIf Dir(cEdoc & "\" & stFolder) = "" Then
         RmDir cEdoc & "\" & stFolder
               
      Else
         sbAddList "失敗,無CSV檔", , True
         Exit Function
      End If
      stFolder = Dir(cEdoc & "\" & cYP & strSrvDate(2) & "*", vbDirectory)
   Loop
   
   '閻啟泰-商標
   stFolder = Dir(cEdoc & "\" & cYT & strSrvDate(2) & "*", vbDirectory)
   Do While stFolder <> ""
      sbAddList "匯入 閻啟泰 商標(" & stFolder & ")"
      stCSV = Dir(cEdoc & "\" & stFolder & "\*.csv")
      If stCSV <> "" Then
         If Import2DB(cEdoc & "\" & stFolder & "\" & stCSV) = True Then
            Name cEdoc & "\" & stFolder As cEdoc & "\" & "BK_" & stFolder '移到備份資料夾以便測試
         Else
            Exit Function
         End If
      'Added by Morgan 2017/11/8 若是空資料夾則刪除後繼續
      ElseIf Dir(cEdoc & "\" & stFolder) = "" Then
         RmDir cEdoc & "\" & stFolder
               
      Else
         sbAddList "失敗,無CSV檔", , True
         Exit Function
      End If
      stFolder = Dir(cEdoc & "\" & cYT & strSrvDate(2) & "*", vbDirectory)
   Loop
   
   '林景郁-專利
   stFolder = Dir(cEdoc & "\" & cLP & strSrvDate(2) & "*", vbDirectory)
   Do While stFolder <> ""
      sbAddList "匯入 林景郁 專利(" & stFolder & ")"
      stCSV = Dir(cEdoc & "\" & stFolder & "\*.csv")
      If stCSV <> "" Then
         If Import2DB(cEdoc & "\" & stFolder & "\" & stCSV) = True Then
            Name cEdoc & "\" & stFolder As cEdoc & "\" & "BK_" & stFolder '移到備份資料夾以便測試
         Else
            Exit Function
         End If
      'Added by Morgan 2017/11/8 若是空資料夾則刪除後繼續
      ElseIf Dir(cEdoc & "\" & stFolder) = "" Then
         RmDir cEdoc & "\" & stFolder
         
      Else
         sbAddList "失敗,無CSV檔", , True
         Exit Function
      End If
      stFolder = Dir(cEdoc & "\" & cLP & strSrvDate(2) & "*", vbDirectory)
   Loop
   
   '林景郁-商標
   stFolder = Dir(cEdoc & "\" & cLT & strSrvDate(2) & "*", vbDirectory)
   Do While stFolder <> ""
      sbAddList "匯入 林景郁 商標(" & stFolder & ")"
      stCSV = Dir(cEdoc & "\" & stFolder & "\*.csv")
      If stCSV <> "" Then
         If Import2DB(cEdoc & "\" & stFolder & "\" & stCSV) = True Then
            Name cEdoc & "\" & stFolder As cEdoc & "\" & "BK_" & stFolder '移到備份資料夾以便測試
         Else
            Exit Function
         End If
      'Added by Morgan 2017/11/8 若是空資料夾則刪除後繼續
      ElseIf Dir(cEdoc & "\" & stFolder) = "" Then
         RmDir cEdoc & "\" & stFolder
      Else
         sbAddList "失敗,無CSV檔", , True
         Exit Function
      End If
      stFolder = Dir(cEdoc & "\" & cLT & strSrvDate(2) & "*", vbDirectory)
   Loop
   
   fnImport = True
   
End Function

Private Function KillFolder(pFolder As String) As Boolean
   
On Error GoTo ErrHnd
   
   If PUB_ChkDir(pFolder) = True Then
      sbAddList "刪除資料夾(" & pFolder & ")"
      ChDir App.path
      'oFileSys.DeleteFile pFolder & "\*.*", True
      'oFileSys.DeleteFolder pFolder & "\*", True
      oFileSys.DeleteFolder pFolder, True
   End If
   KillFolder = True
   Exit Function
ErrHnd:
   sbAddList "" & Err.Number & "," & Err.Description
   
End Function

Private Function Import2DB(pCSV As String) As Boolean
   Dim strText As String
   Dim arrRow() As String
   Dim arrCell() As String, arrCell2() As String, arrED02() As String, arrED15() As String, arrED16() As String, arrED24() As String
   Dim idx1 As Integer, idx2 As Integer, idx3 As Integer, iNo As Integer
   Dim stSQL As String, stValues As String, intR As Integer
   Dim stED01 As String, stED02 As String, stED10 As String, stED16 As String, stED23 As String, stED01A As String
   Dim stFolder As String, stFileName As String, stFileList As String, stSrcFileName As String
   Dim arrEDidx(31) As Integer
   Dim arrColNames() As String
   Dim strNewCol As String
   Dim iRecs As Integer
   Dim stNoRecIdList As String '不需收文清單
   Dim stAddFolder As String, stAddFiles As String, stAddFileList As String
   Dim stJoinIdList As String '檔案合併清單
   Dim stDupNoList As String, stDownFile As String
   Dim stTCtlCaseList As String 'T與FCT共同管控案件註冊號清單 Added by Morgan 2022/1/14
   
   Const cDelimiter As String = """,""" '欄位區隔符號(名稱)
   Const cDelimiter2 As String = """," '欄位區隔符號(值) ' 重新下載最後面的欄位值沒有雙引號(")故值改用(",)區隔(取值時去除頭尾的雙引號)
   
   Const BlockSize = 500000
   
   
On Error GoTo ErrHnd

   stTCtlCaseList = Pub_GetSpecMan("T與FCT共同管控案件") 'Added by Morgan 2022/1/14
   
   strText = ReadTextFile(pCSV)
   '案由有逗號(,)與欄位區隔符號重複,雙引號不可先去除
   '首列(欄位名)與資料列會用 chr(9) & chr(10) 分隔
   '資料列又改為都用 chr(10) 分隔
   strText = Replace(strText, Chr(9), "")
   strText = Replace(strText, Chr(13) & Chr(10), Chr(10))
   arrRow = Split(strText, Chr(10))
   
   arrCell = Split(arrRow(LBound(arrRow)), cDelimiter)
   ReDim arrColNames(UBound(arrCell)) As String
   
   strNewCol = ""
   arrEDidx(9) = -1
   For idx1 = LBound(arrCell) To UBound(arrCell)
      strText = GetStr(arrCell(idx1))
      Select Case strText
      Case "發文文號"
         arrColNames(idx1) = "ED01": arrEDidx(1) = idx1
         
      Case "原申請案號"
         arrColNames(idx1) = "ED02": arrEDidx(2) = idx1
         
      Case "送達時間"
         arrColNames(idx1) = "ED03": arrEDidx(3) = idx1
         
      Case "案由"
         arrColNames(idx1) = "ED04": arrEDidx(4) = idx1
         
      Case "簽收時間", "補送簽收時間"
         arrColNames(idx1) = "ED05": arrEDidx(5) = idx1
         
      Case "簽收人"
         arrColNames(idx1) = "ED06": arrEDidx(6) = idx1
         
      Case "受送達人"
         arrColNames(idx1) = "ED07": arrEDidx(7) = idx1
         
      Case "發文日期"
         arrColNames(idx1) = "ED08": arrEDidx(8) = idx1
         
      Case "檔案"
         arrColNames(idx1) = "ED09": arrEDidx(9) = idx1
         
      Case "案件種類"
         arrColNames(idx1) = "ED10": arrEDidx(10) = idx1
         
      Case "案號類別"
         arrColNames(idx1) = "ED15": arrEDidx(15) = idx1
         
      Case "案號"
         arrColNames(idx1) = "ED16": arrEDidx(16) = idx1
         
      'Added by Morgan 2014/4/14
      Case "發文字號"
         arrColNames(idx1) = "ED17": arrEDidx(17) = idx1
         
      Case "處理期限"
         arrColNames(idx1) = "ED18": arrEDidx(18) = idx1
         
      Case "處理期間"
         arrColNames(idx1) = "ED19": arrEDidx(19) = idx1
      Case "受文者序號"
         arrColNames(idx1) = "ED23": arrEDidx(23) = idx1
      Case "相關案號"
         arrColNames(idx1) = "ED24": arrEDidx(24) = idx1
      Case "承審委員"
         arrColNames(idx1) = "ED25": arrEDidx(25) = idx1
      Case "IPC分類"
         arrColNames(idx1) = "ED26": arrEDidx(26) = idx1
      Case "事務所案號"
         arrColNames(idx1) = "ED27": arrEDidx(27) = idx1
      Case "正副本"
         arrColNames(idx1) = "ED28": arrEDidx(28) = idx1
      Case "補送下載時間"
         arrColNames(idx1) = "ED29": arrEDidx(29) = idx1
      Case "受文者"
         arrColNames(idx1) = "ED31": arrEDidx(31) = idx1
      Case "檔案路徑" 'Added by Morgan 2017/6/15 不必匯入系統
         arrColNames(idx1) = "ED22": arrEDidx(22) = idx1 'Added by Morgan 2017/7/28 改要存
      Case Else
         strNewCol = strNewCol & strText & vbCrLf
      End Select
   Next
   
   If UBound(arrRow) = LBound(arrRow) Then
      sbAddList "失敗,無資料(" & arrCell(0) & ")", , True
      Exit Function
   End If

   '先將 pdf 合併
   For idx1 = LBound(arrRow) + 1 To UBound(arrRow)
      'Modified by Morgan 2017/5/25 不收文的不必合併
      If arrRow(idx1) <> "" And InStr(stNoRecIdList, "(" & idx1 & ")") = 0 Then
         arrCell = Split(arrRow(idx1), cDelimiter2)
         stAddFolder = ""
         stAddFiles = ""
         stAddFileList = ""
         '檢查後面資料若有發文號重複者後面的預設不收文,若附件不同時複製不同附件到前者的資料夾
         For idx2 = idx1 + 1 To UBound(arrRow)
            If arrRow(idx2) <> "" Then
               arrCell2 = Split(arrRow(idx2), cDelimiter2)
               '發文號相同
               If arrCell2(arrEDidx(1)) = arrCell(arrEDidx(1)) Then
                  '附件不同
                  If arrCell2(arrEDidx(9)) <> arrCell(arrEDidx(9)) Then
                     'Added by Morgan 2017/5/25 考慮資料作業失敗但已合併情形
                     'Modified by Morgan 2017/7/28 改抓"檔案路徑"欄位
                     'stFolder = GetStr(arrCell(arrEDidx(2))) & "_" & GetStr(arrCell(arrEDidx(1))) & "_" & GetStr(arrCell(arrEDidx(23)))
                     stFolder = GetStr(arrCell(arrEDidx(22)))
                     stFolder = Left(pCSV, InStrRev(pCSV, "\")) & stFolder
                     stFileName = "$" & GetStr(arrCell(arrEDidx(1))) & ".pdf"
                     '合併檔已存在時直接加清單
                     If oFileSys.FileExists(stFolder & "\" & stFileName) = True Then
                        '檔案已合併索引清單
                        stJoinIdList = stJoinIdList & "(" & idx1 & ")"
                     Else
                        'stAddFolder = GetStr(arrCell2(arrEDidx(2))) & "_" & GetStr(arrCell2(arrEDidx(1))) & "_" & GetStr(arrCell2(arrEDidx(23)))
                        stAddFolder = GetStr(arrCell2(arrEDidx(22)))
                        stAddFolder = Left(pCSV, InStrRev(pCSV, "\")) & stAddFolder
                        stAddFiles = GetStr(arrCell2(arrEDidx(9)))
                        'Modified by Morgan 2017/7/28 改抓"檔案路徑"欄位
                        'stFolder = GetStr(arrCell(arrEDidx(2))) & "_" & GetStr(arrCell(arrEDidx(1))) & "_" & GetStr(arrCell(arrEDidx(23)))
                        stFolder = GetStr(arrCell(arrEDidx(22)))
                        stFolder = Left(pCSV, InStrRev(pCSV, "\")) & stFolder
                        If CopyFile(stAddFolder, stFolder, stAddFiles) = True Then
                           '檔案已合併索引清單
                           stJoinIdList = stJoinIdList & "(" & idx1 & ")"
                           
                           stAddFileList = stAddFileList & IIf(stAddFileList <> "", ";", "") & stAddFiles
                        Else
                           Exit Function
                        End If
                        
                     End If
                     'End If
                  End If
                  
                  '不收文的索引清單
                  stNoRecIdList = stNoRecIdList & "(" & idx2 & ")"
               End If
            End If
         Next
         
         '檢查卷宗區是否已有該發文號,若有則下載檔案與目前檔案合併
         stED01 = GetStr(arrCell(arrEDidx(1)))
         'Modified by Morgan 2019/8/15 要加檔名的條件，因為合併前的檔案也會放卷宗區。Ex:發文號:10890788710
         'cnnConnection.Execute "update casepaperpdf set cpp02=cpp02 where cpp01='" & stED01 & "'", intR
         cnnConnection.Execute "update casepaperpdf set cpp02=cpp02 where cpp01='" & stED01 & "' and cpp02='$'||cpp01||'.pdf'", intR
         If intR = 1 Then
            'Modified by Morgan 2017/7/28 改抓"檔案路徑"欄位
            'stFolder = GetStr(arrCell(arrEDidx(2))) & "_" & GetStr(arrCell(arrEDidx(1))) & "_" & GetStr(arrCell(arrEDidx(23)))
            stFolder = GetStr(arrCell(arrEDidx(22)))
            stFolder = Left(pCSV, InStrRev(pCSV, "\")) & stFolder
            stFileName = "$" & GetStr(arrCell(arrEDidx(1))) & ".pdf"
            If oFileSys.FileExists(stFolder & "\" & stFileName) = False Then
               stDownFile = "$" & stED01 & ".pdf"
               If Dir(App.path & "\" & stDownFile) <> "" Then Kill App.path & "\" & stDownFile
               If PUB_GetAttachFile_CPP(stED01, stDownFile, App.path) = True Then
                  If oFileSys.FileExists(stDownFile) = True Then
                     Set oFile = oFileSys.GetFile(stDownFile)
                     stAddFiles = stFolder & "\$$" & stED01 & ".pdf"
                     If Dir(stAddFiles) <> "" Then Kill stAddFiles
                     oFile.Move stAddFiles
                     stAddFileList = stAddFileList & IIf(stAddFileList <> "", ";", "") & stAddFiles
                  End If
               End If
            End If
            stDupNoList = stDupNoList & IIf(stDupNoList <> "", ",", "") & stED01
         End If
         
         strText = GetStr(arrCell(arrEDidx(9)))
         '將相同發文號的附件字串合併並將正本放前面
         If stAddFileList <> "" Then
            If GetStr(arrCell(arrEDidx(28))) = "正本" Then
               strText = strText & ";" & stAddFileList
            Else
               strText = stAddFileList & ";" & strText
            End If
         End If
         
         '合併附件
         If InStr(LCase(strText), ".pdf;") > 0 Then
            'pdf路徑
            'Modified by Morgan 2017/7/28 改抓"檔案路徑"欄位
            'stFolder = GetStr(arrCell(arrEDidx(2))) & "_" & GetStr(arrCell(arrEDidx(1))) & "_" & GetStr(arrCell(arrEDidx(23)))
            stFolder = GetStr(arrCell(arrEDidx(22)))
            stFolder = Left(pCSV, InStrRev(pCSV, "\")) & stFolder
            '合併後檔名=$發文文號.pdf
            stFileName = "$" & GetStr(arrCell(arrEDidx(1))) & ".pdf"
            If oFileSys.FileExists(stFolder & "\" & stFileName) = False Then '若檔案已存在則跳過(當程式無法合併時,由人工放入合併檔)
               strExc(1) = GetStr(arrCell(arrEDidx(2))) '申請號
               strExc(2) = GetStr(arrCell(arrEDidx(10))) '案件種類
               strExc(3) = GetStr(arrCell(arrEDidx(4))) '案由
               strExc(4) = GetStr(arrCell(arrEDidx(1)))
               If JoinPdf(strText, stFileName, stFolder, strExc(1), strExc(2), , strExc(3), strExc(4)) = False Then
                  Exit Function
               End If
            End If
         End If

      End If
   Next
   
   cnnConnection.BeginTrans
   
On Error GoTo ErrHndT
   
   For idx1 = LBound(arrRow) + 1 To UBound(arrRow)
      Debug.Print idx1 & "/" & UBound(arrRow)
      If arrRow(idx1) <> "" Then
         stFolder = ""
         stSQL = "Insert into EDocument("
         stValues = ""
         stED01 = ""
         stED02 = ""
         stED10 = ""
         stED16 = ""
         stED23 = ""
         arrCell = Split(arrRow(idx1), cDelimiter2)
         For idx2 = LBound(arrColNames) To UBound(arrColNames)
            If arrColNames(idx2) <> "" And idx2 <= UBound(arrCell) Then
               arrCell(idx2) = GetStr(arrCell(idx2)) '去除多餘的符號
               strText = arrCell(idx2)
               If stValues <> "" Then
                  stSQL = stSQL & ","
                  stValues = stValues & ","
               End If
               '簽收,送達日期時間
               If arrColNames(idx2) = "ED03" Or arrColNames(idx2) = "ED05" Or arrColNames(idx2) = "ED21" Or arrColNames(idx2) = "ED29" Then
                  '民國年改西元年 Ex. 103/8/27 11:58
                  If InStr(strText, "/") = 4 Then
                     strText = (Val(Left(strText, 3)) + 1911) & Mid(strText, 4)
                  End If
                  'end 2014/9/29
                  stValues = stValues & "to_date('" & strText & "','yyyy/mm/dd hh24:mi:ss')"
               '發文日期
               ElseIf arrColNames(idx2) = "ED08" Then
                  stValues = stValues & (Val(Replace(strText, "/", "")) + 19110000)
               'Added by Morgan 2020/8/11
               '事務所案號:不要匯入,正常不會給,有也不一定正確(智慧局周小姐說沒有維護也不能給)
               ElseIf arrColNames(idx2) = "ED27" Then
                  stValues = stValues & "''"
               'end 2020/8/11
               Else
                  '處理期限
                  If arrColNames(idx2) = "ED18" Then
                     strText = Replace(strText, "/", "")
                     
                  '下列欄位含;號時只記錄第1筆,此1文多案情形,目前第2案以後人工影印後以紙本方式輸入 ex.106/5/25 發文號:10680268560
                  '申請案號
                  ElseIf arrColNames(idx2) = "ED02" And InStr(strText, ";") > 0 Then
                     'Added by Morgan 2023/5/22
                     stSQL = stSQL & "ED33,"
                     stValues = stValues & "'" & ChgSQL(strText) & "'" & ","
                     'end 2023/5/22
                     strText = Left(strText, InStr(strText, ";") - 1)
                     stED02 = strText
                  '案號類別
                  ElseIf arrColNames(idx2) = "ED15" And InStr(strText, ";") > 0 Then
                     'Added by Morgan 2023/5/22
                     stSQL = stSQL & "ED34,"
                     stValues = stValues & "'" & ChgSQL(strText) & "'" & ","
                     'end 2023/5/22
                     strText = Left(strText, InStr(strText, ";") - 1)
                  '案號
                  ElseIf arrColNames(idx2) = "ED16" And InStr(strText, ";") > 0 Then
                     'Added by Morgan 2023/5/22
                     stSQL = stSQL & "ED35,"
                     stValues = stValues & "'" & ChgSQL(strText) & "'" & ","
                     'end 2023/5/22
                     strText = Left(strText, InStr(strText, ";") - 1)
                     stED16 = strText
                  End If
                  stValues = stValues & "'" & ChgSQL(strText) & "'"
               End If
               stSQL = stSQL & arrColNames(idx2)
            End If
         Next
         
         stSQL = stSQL & ") values (" & stValues & ")"
         cnnConnection.Execute stSQL, intR
         
         stED01 = arrCell(arrEDidx(1))
         If stED02 = "" Then stED02 = arrCell(arrEDidx(2))
         stED10 = arrCell(arrEDidx(10))
         If stED16 = "" Then stED16 = arrCell(arrEDidx(16))
         stED23 = arrCell(arrEDidx(23))
         
         UpdateOurRef stED01, stED02, stED10, stED16, stED23 '更新事務所案號(ED27)
         
         iRecs = iRecs + 1

         If InStr(stNoRecIdList, "(" & idx1 & ")") = 0 Then '要收文的才要上傳檔案
         
            'Added by Morgan 2017/5/26 發文號重複的要刪除
            If InStr(stDupNoList, stED01) > 0 Then
               cnnConnection.Execute "delete casepaperpdf where cpp01='" & stED01 & "'", intR
            End If
            'end 2017/5/26
            
            '附件資料夾名稱 申請號_發文號_受文者序號
            'Modified by Morgan 2017/7/28 改抓"檔案路徑"欄位
            'stFolder = arrCell(arrEDidx(2)) & "_" & arrCell(arrEDidx(1)) & "_" & arrCell(arrEDidx(23))
            stFolder = GetStr(arrCell(arrEDidx(22)))
            '上傳檔案
            stFolder = Left(pCSV, InStrRev(pCSV, "\")) & stFolder
            
            '單檔改直接上傳原始檔(不再產生合併檔)
            stFileName = arrCell(arrEDidx(9))
            If InStr(arrCell(arrEDidx(9)), ".pdf;") > 0 Or InStr(stJoinIdList, "(" & idx1 & ")") > 0 Or InStr(stDupNoList, stED01) > 0 Then
               stFileName = "$" & stED01 & ".pdf"
               'Added by Morgan 2017/6/2
               '若為合併檔,原始檔案也要上傳列印時改印原始檔(當附件為基數頁雙面列印會將第1份附件的最後1頁與第2份附件的第1頁印在同一張上,如為正副本或繳費單時要另外影印才能寄客戶)
               stSrcFileName = Dir(stFolder & "\" & stFileName & ".*")
               Do While stSrcFileName <> ""
                  If stSrcFileName <> stFileName Then
                     Set oFile = oFileSys.GetFile(stFolder & "\" & stSrcFileName)
                     SaveAttFile_PDF stED01, stFolder & "\" & stSrcFileName, stSrcFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True, , , , , , "0"
                  End If
                  stSrcFileName = Dir()
               Loop
               'end 2017/6/2
               
               'Added by Morgan 2023/1/12
               '證書/註冊證
               stSrcFileName = Dir(stFolder & "\$" & stED01 & ".CERT.pdf")
               If stSrcFileName <> "" Then
                  Set oFile = oFileSys.GetFile(stFolder & "\" & stSrcFileName)
                  SaveAttFile_PDF stED01, stFolder & "\" & stSrcFileName, stSrcFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True, , , , , , "0"
               End If
               '若有其他附件(正常不該有)
               stSrcFileName = Dir(stFolder & "\$" & stED01 & ".CERT.?.pdf")
               Do While stSrcFileName <> ""
                  Set oFile = oFileSys.GetFile(stFolder & "\" & stSrcFileName)
                  SaveAttFile_PDF stED01, stFolder & "\" & stSrcFileName, stSrcFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True, , , , , , "0"
                  stSrcFileName = Dir()
               Loop
               'end 2023/1/12
            End If
            Set oFile = oFileSys.GetFile(stFolder & "\" & stFileName)
            SaveAttFile_PDF stED01, stFolder & "\" & stFileName, "$" & stED01 & ".pdf", Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True, , , , , , "0"
         End If
         
         'Added by Morgan 2021/6/16
         '檢查商標註冊號是否為 01922108、01922109 (T,FCT都要通知)
         'Modified by Morgan 2022/1/14 增加案件,改抓系統特殊設定
         'If stED10 = "T" And (stED16 = "01922108" Or stED16 = "01922109") Then
         If stED10 = "T" And InStr(";" & stTCtlCaseList & ";", ";" & stED16 & ";") > 0 Then
         'end 2022/1/14
            strExc(0) = "select tm01,tm01||tm02||tm03||tm04 CNo from edocument a,trademark where ed01='" & stED01 & "' and ed23='" & stED23 & "' and tm15(+)=ed16 and tm01 in('T','FCT')  and tm10='000' and tm57 is null and tm01||tm02||tm03||tm04<>ed27"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               stED01A = stED01 & RsTemp.Fields("tm01")
               stSQL = ""
               stValues = ""
               For idx2 = LBound(arrColNames) To UBound(arrColNames)
                  If arrColNames(idx2) <> "" Then
                     If stSQL <> "" Then
                        stSQL = stSQL & ","
                        stValues = stValues & ","
                     End If
                     stSQL = stSQL & arrColNames(idx2)
                     
                     If arrColNames(idx2) = "ED01" Then
                        stValues = stValues & "'" & stED01A & "'"
                     ElseIf arrColNames(idx2) = "ED27" Then
                        stValues = stValues & "'" & RsTemp.Fields("CNo") & "'"
                     Else
                        stValues = stValues & arrColNames(idx2)
                     End If
                  End If
               Next
               stSQL = "Insert into EDocument(" & stSQL & ") select " & stValues & " from edocument where ed01='" & stED01 & "' and ed23='" & stED23 & "'"
               cnnConnection.Execute stSQL, intR
   
               Set oFile = oFileSys.GetFile(stFolder & "\" & stFileName)
               SaveAttFile_PDF stED01A, stFolder & "\" & stFileName, "$" & stED01A & ".pdf", Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True, , , , , , "0"
            End If
         End If
         'end 2021/6/16
        
      'Added by Morgan 2021/11/8 商標一文多案通知補正只要產生第一件之電子公文--陳金蓮,林桂英
      'Modified by Morgan 2023/5/22 排除副本(非本所辦理的案件)，案由增加移轉案通知補正(改抓後4碼)，本所案號增加主案判斷(發文號=收文號)
      'If (stED10 = "T" And arrCell(arrEDidx(4)) = "通知補正") Then
      If (stED10 = "T" And Right(arrCell(arrEDidx(4)), 4) = "通知補正") And arrCell(arrEDidx(28)) <> "副本" Then
         If InStr(arrCell(arrEDidx(2)), ";") > 0 Or (stED16 <> "" And InStr(arrCell(arrEDidx(24)), ";") > 0) Then
            stSQL = "update edocument set ed21='商標一文多案通知補正;'||ed21 where ed01='" & stED01 & "' and ed23='" & stED23 & "'"
            cnnConnection.Execute stSQL, intR
            
            UpdMultiDoc stED01, stED23 'Added by Morgan 2023/5/22
         End If
      Else
      'end 2021/11/8
            
         'Added by Morgan 2017/6/6
         '1文多案
         If InStr(arrCell(arrEDidx(2)), ";") > 0 Then
            'Added by Morgan 2023/12/6 正本副本皆有時除外 Ex:ED01=1129124497
            intR = 0
            If arrCell(arrEDidx(28)) = "副本" Then
               stSQL = "update edocument set ed21=ed21 where ed01='" & stED01 & "' and ed28='正本'"
               cnnConnection.Execute stSQL, intR
            End If
            If intR = 0 Then
            'end 2023/12/6
            
               arrED02 = Split(arrCell(arrEDidx(2)), ";")
               arrED15 = Split(arrCell(arrEDidx(15)), ";")
               arrED16 = Split(arrCell(arrEDidx(16)), ";")
               iNo = 0
               For idx3 = LBound(arrED02) To UBound(arrED02)
                  If arrED02(idx3) <> "" And arrED02(idx3) <> stED02 Then
                     iNo = iNo + 1
                     stED01A = stED01 & Format(iNo, "000")
                     stSQL = ""
                     stValues = ""
                     For idx2 = LBound(arrColNames) To UBound(arrColNames)
                        If arrColNames(idx2) <> "" Then
                           If stSQL <> "" Then
                              stSQL = stSQL & ","
                              stValues = stValues & ","
                           End If
                           stSQL = stSQL & arrColNames(idx2)
                           
                           If arrColNames(idx2) = "ED01" Then
                              stValues = stValues & "'" & stED01A & "'"
                           ElseIf arrColNames(idx2) = "ED02" Then
                              stValues = stValues & "'" & arrED02(idx3) & "'"
                              stED02 = arrED02(idx3)
                           ElseIf arrColNames(idx2) = "ED15" Then
                              stValues = stValues & "'" & arrED15(idx3) & "'"
                           ElseIf arrColNames(idx2) = "ED16" Then
                              stValues = stValues & "'" & arrED16(idx3) & "'"
                              stED16 = arrED16(idx3)
                           ElseIf arrColNames(idx2) = "ED27" Then
                              stValues = stValues & "''"
                           Else
                              stValues = stValues & arrColNames(idx2)
                           End If
                        End If
                     Next
                     stSQL = "Insert into EDocument(" & stSQL & ") select " & stValues & " from edocument where ed01='" & stED01 & "' and ed23='" & stED23 & "'"
                     cnnConnection.Execute stSQL, intR
                     
                     UpdateOurRef stED01A, stED02, stED10, stED16, stED23 '更新事務所案號(ED27)
                     
                     iRecs = iRecs + 1
                     
                     If InStr(stNoRecIdList, "(" & idx1 & ")") = 0 Then '要收文的才要上傳檔案
                        'Added by Morgan 2017/6/22 發文號重複的要刪除
                        If InStr(stDupNoList, stED01) > 0 Then
                           cnnConnection.Execute "delete casepaperpdf where cpp01='" & stED01A & "'", intR
                        End If
                        'end 2017/6/22
                        Set oFile = oFileSys.GetFile(stFolder & "\" & stFileName)
                        SaveAttFile_PDF stED01A, stFolder & "\" & stFileName, "$" & stED01A & ".pdf", Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True, , , , , , "0"
                     End If
                     
                     If iNo = 999 Then Exit For
                  End If
               Next
            
            End If 'Added by Morgan 2023/12/6
            
         'Added by Morgan 2017/6/19
         '一文多案第2案以後的案號改放在相關案號
         'Memo by Morgan 2021/11/8 目前看來這段好像沒作用，因為申請案號(ED02)還是會放多案號，相關案號(ED24)反而沒放，不確定是否跟案件數有關
         ElseIf stED10 = "T" And stED16 <> "" And InStr(arrCell(arrEDidx(24)), ";") > 0 Then
            
            arrED24 = Split(arrCell(arrEDidx(24)), ";")
            iNo = 0
            For idx3 = LBound(arrED24) To UBound(arrED24)
               If arrED24(idx3) <> "" And arrED24(idx3) <> stED16 Then
                  iNo = iNo + 1
                  stED01A = stED01 & Format(iNo, "000")
                  stSQL = ""
                  stValues = ""
                  For idx2 = LBound(arrColNames) To UBound(arrColNames)
                     If arrColNames(idx2) <> "" Then
                        If stSQL <> "" Then
                           stSQL = stSQL & ","
                           stValues = stValues & ","
                        End If
                        stSQL = stSQL & arrColNames(idx2)
                        
                        If arrColNames(idx2) = "ED01" Then
                           stValues = stValues & "'" & stED01A & "'"
                        
                        '若案號類別不是註冊號時,原申請案號欄位放相關案號
                        ElseIf arrColNames(idx2) = "ED02" And arrCell(arrEDidx(15)) <> "註冊號" Then
                           stValues = stValues & "'" & arrED24(idx3) & "'"
                           stED02 = arrED24(idx3)
                           
                        ElseIf arrColNames(idx2) = "ED16" Then
                           stValues = stValues & "'" & arrED24(idx3) & "'"
                           stED16 = arrED24(idx3)
                           
                        ElseIf arrColNames(idx2) = "ED27" Then
                           stValues = stValues & "''"
                           
                        Else
                           stValues = stValues & arrColNames(idx2)
                        End If
                     End If
                  Next
                  stSQL = "Insert into EDocument(" & stSQL & ") select " & stValues & " from edocument where ed01='" & stED01 & "' and ed23='" & stED23 & "'"
                  cnnConnection.Execute stSQL, intR
                  
                  UpdateOurRef stED01A, stED02, stED10, stED16, stED23 '更新事務所案號(ED27)
                  
                  iRecs = iRecs + 1
                  
                  If InStr(stNoRecIdList, "(" & idx1 & ")") = 0 Then '要收文的才要上傳檔案
                     'Added by Morgan 2017/6/22 發文號重複的要刪除
                     If InStr(stDupNoList, stED01) > 0 Then
                        cnnConnection.Execute "delete casepaperpdf where cpp01='" & stED01A & "'", intR
                     End If
                     'end 2017/6/22
                     Set oFile = oFileSys.GetFile(stFolder & "\" & stFileName)
                     SaveAttFile_PDF stED01A, stFolder & "\" & stFileName, "$" & stED01A & ".pdf", Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "U", True, , , , , , "0"
                  End If
                  
                  If iNo = 999 Then Exit For
               End If
            Next
         End If
         'end 2017/6/6
         
      End If 'Added by Morgan 2021/11/8
         
         'Modified by Morgan 2017/6/27
         '目前發生1.正副本 2.多正本 3.本所非第1受文者
         '改受文者序號最小的設要收文
         For iNo = 1 To 10
            cnnConnection.Execute "update edocument set ed20='' where ed01>='" & stED01 & "' and ed01<='" & stED01 & "999' and ed23='" & iNo & "'", intR
            If intR > 0 Then
               cnnConnection.Execute "update edocument set ed20='N' where ed01>='" & stED01 & "' and ed01<='" & stED01 & "999' and ed23<>'" & iNo & "'", intR
               Exit For
            End If
         Next
         
      End If
   Next
   
   cnnConnection.CommitTrans
   
   sbAddList "完成,共 " & iRecs & " 筆" & IIf(strNewCol <> "", "(CSV有新欄位 " & strNewCol & " 未匯入)", ""), , True
   
   Import2DB = True
   Exit Function
   
ErrHndT:
   cnnConnection.RollbackTrans
   
ErrHnd:
   If Err.Number <> 0 Then
      sbAddList "" & Err.Number & "," & Err.Description & "(發文文號:" & stED01 & ")"
   End If
   
End Function

Private Sub cmbPrinter_Click()
   'Added by Morgan 2018/5/7
   If Me.cmbPrinter.Tag <> "" And Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
      MsgBox "您點選的印表機與預設不同，若為測試請在程式結束前改回以免影響自動列印作業！", vbExclamation
   End If
End Sub

Private Function ChkAutoRun(pReason As String) As Boolean
   strExc(0) = "select * from workday where wd01=" & strSrvDate(1)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp("wd07") = "Y" Then
         pReason = "設定手動下載"
      Else
         ChkAutoRun = True
      End If
   Else
      pReason = "非工作日"
   End If
End Function

Private Sub cmdAutoRun_Click()
   
   'Modified by Morgan 2019/9/6
   '自動執行時增加檢查是否有設定手動下載
   If Me.ActiveControl <> cmdAutoRun Then
      If ChkAutoRun(strExc(1)) = False Then
         sbAddList strExc(1) & "，自動作業取消！"
         cmdExit.Value = True
         Exit Sub
      End If
      
   ElseIf ChkWorkDay(strSrvDate(1)) = False Then
      sbAddList "非工作日，自動作業取消！"
      Exit Sub
      
   End If
   'end 2019/9/6
      
   If UCase(pub_DbTerminalName) <> UCase(正式資料庫電腦名稱) Then
      sbAddList "目前連線為測試資料庫不可自動作業！"
      
   Else
      m_bolContinue = True
      '下載
      Command1(0).Value = True
      '匯入
      If m_bolContinue Then
         Command1(1).Value = True
      End If
      
      'm_bolContinue = False 'Added by Morgan 2019/12/27 自動列印有問題,先改手動執行
      
      '清單及公文
      If m_bolContinue Then
         Command1(3).Value = True
      End If
      'EMail
      If m_bolContinue Then
         Check1.Value = vbChecked
         Command1(2).Value = True
         
         If m_bolContinue And m_bolAutoUnload Then cmdExit.Value = True 'Added by Morgan 2017/12/5
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
   Timer1.Enabled = False
   fraCountDown.Visible = False
End Sub

Private Sub cmdExit_Click()
   If App.EXEName = "teAutoEDoc" Then
      Unload Me
      End
   Else
      Unload Me
   End If
End Sub


Private Sub Command1_Click(Index As Integer)
   Dim bol
   Me.Enabled = False
   Screen.MousePointer = vbHourglass
   Select Case Index
   Case 0 '下載
      sbAddList "公文下載...開始", True
      If fnDowload() = True Then
         sbAddList "公文下載...結束"
      Else
         m_bolContinue = False
      End If
      
   Case 1 '匯入
      sbAddList "公文匯入...開始", True
      If fnImport() = True Then
         sbAddList "公文匯入...結束"
      Else
         m_bolContinue = False
      End If
      
   Case 2 'EMail
      sbAddList "EMail通知...開始", True
      If fnEMail() = True Then
         sbAddList "EMail通知...結束"
      Else
         m_bolContinue = False
      End If
      
   Case 3 '列印清單及附件
      m_iPCount = 0
      m_iLCount = 0 'Added by Morgan 2019/6/18
      sbAddList "列印清單及附件...開始", True
      If fnReport(True) Then
         sbAddList "列印清單及附件...結束(" & m_iLCount & " + " & m_iPCount & " = " & m_iLCount + m_iPCount & ")"
      Else
         m_bolContinue = False
      End If
      
   Case 4 '列印清單
      m_iLCount = 0 'Added by Morgan 2019/6/18
      sbAddList "列印清單...開始", True
      If fnReport(, True) = True Then
         sbAddList "列印清單...結束(" & m_iLCount & ")"
      End If
      
   End Select
   Screen.MousePointer = vbDefault
   Me.Enabled = True
End Sub

Private Function fnReport(Optional pPrintAtt As Boolean = False, Optional pReport As Boolean) As Boolean
   Dim strRptDate As String, strPrinter As String, intOrientation As Integer
   Dim process_id As Long
   Dim process_handle As Long
   
   If txtDate = "" Then
      MsgBox "請輸入簽收日期！", vbExclamation
   Else
      strRptDate = TransDate(txtDate, 1)
      If ChkDate(strRptDate) = True Then
      
         strPrinter = Printer.DeviceName
         intOrientation = Printer.Orientation
         PUB_RestorePrinter cmbPrinter
         
         If pPrintAtt Then
            m_PdfReader = PUB_SetFileAssociation
            process_id = Shell("""" & m_PdfReader & """", vbNormalNoFocus)
            process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
            
            Sleep 5000 'Added by Morgan 2019/12/24 最近常當掉,加5秒等待,看看是否因初次載入時間較長造成
         End If
         
         'Removed by Morgan 2025/2/11 P 改清單公文都不印
         'Modified by Morgan 2025/9/24 只印清單時還是要能印
         If pReport Then
            If Combo1 = "" Or Combo1 = "P" Then
               ReportP strRptDate, pPrintAtt
            End If
         End If
         'end 2025/2/11
         
         'Removed by Morgan 2023/7/3 FCP 改清單公文都不印
         'Modified by Morgan 2025/9/24 只印清單時還是要能印
         If pReport Then
            If Combo1 = "" Or Combo1 = "FCP" Then
               ReportFCP strRptDate, pPrintAtt
            End If
         End If
         'end 2023/7/3
         
         If Combo1 = "" Or Combo1 = "T" Then
            ReportT strRptDate, pPrintAtt
         End If
         If Combo1 = "" Or Combo1 = "FCT" Then
            ReportFCT strRptDate, pPrintAtt
         End If
         
         If process_handle <> 0 Then
            TerminateProcess process_handle, 0&
            CloseHandle process_handle
         End If
         
         fnReport = True
         PUB_RestorePrinter strPrinter, intOrientation
      End If
   End If
End Function

Private Sub SetFrame()
   With fraCountDown
      .Left = 0
      .Top = -100
      .Width = Me.Width - 100
      .Height = Me.Height - 250
      lblCountDown.Top = .Height / 2 - lblCountDown.Height / 2
      lblCountDown.Left = .Width / 2 - lblCountDown.Width / 2
      
      
      lblCount.Left = .Width / 2 - lblCount.Width / 2
      lblCount.Top = lblCountDown.Top - lblCount.Height
      
      cmdCancel.Left = .Width / 2 - cmdCancel.Width / 2
      cmdCancel.Top = lblCountDown.Top + lblCountDown.Height
   End With
End Sub

Private Sub UpdateED27(pED13 As String)

strExc(0) = "select * From edocument where ed13=" & pED13
intI = 1
Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
If intI = 1 Then
With RsTemp
Do While Not .EOF
   'Modified by Morgan 2022/6/29 +"" & .Fields("ed15")
   UpdateOurRef .Fields("ed01"), .Fields("ed02"), .Fields("ed10"), "" & .Fields("ed16"), .Fields("ed23"), "" & .Fields("ed15")
   .MoveNext
Loop
End With
End If
End Sub

Private Sub Command2_Click()
   If txtDate <> "" Then UpdateED27 DBDATE(txtDate)
End Sub

Private Sub Form_Activate()
   Static bolActived As Boolean
   
   If bolActived = False Then
      SetFrame
      Timer1.Enabled = True
      Timer1.Interval = 1000
      lblCountDown = 5
      bolActived = True
      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Then
         cmdCancel.Value = True
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   PUB_SetPrinter pub_HostName & "-" & Me.Name, cmbPrinter
   
   PUB_SetPrinter pub_HostName & "-" & Me.Name & "-2", cmbPrinter2
   cPFeeForm = "\\" & strPat1Path & "\Fee_Form" 'Added by Lydia 2024/07/22 P案專利申請書存放路徑
   lstHistory.Clear
   txtExe.Text = cExePath & "\" & cExeName
   txtDate.Text = strSrvDate(2)
   txtOutput(0) = cEdoc & "\" & cYP & strSrvDate(2) & "######"
   txtOutput(1) = cEdoc & "\" & cYT & strSrvDate(2) & "######"
   txtOutput(2) = cEdoc & "\" & cLP & strSrvDate(2) & "######"
   txtOutput(3) = cEdoc & "\" & cLT & strSrvDate(2) & "######"
   
   Combo1.Clear
   Combo1.AddItem ""
   Combo1.AddItem "P"
   Combo1.AddItem "FCP"
   Combo1.AddItem "T"
   Combo1.AddItem "FCT"
   
   Timer2.Interval = 1000
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   'Added by Morgan 2018/5/7
   'Modified by Morgan 2021/7/21 排除被呼叫列印清單
   If m_bCalled = False And ((Me.cmbPrinter.Tag <> "" And Me.cmbPrinter.Text <> Me.cmbPrinter.Tag) Or (Me.cmbPrinter2.Tag <> "" And Me.cmbPrinter2.Text <> Me.cmbPrinter2.Tag)) Then
      If MsgBox("您已變更預設印表機！該設定將影響下次自動列印作業，是否確定要繼續？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         Cancel = 1
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   If Me.cmbPrinter.Text <> Me.cmbPrinter.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, pub_HostName & "-" & Me.Name, Me.cmbPrinter.Name, "0", "0", Me.cmbPrinter.Text
   End If
   If Me.cmbPrinter2.Text <> Me.cmbPrinter2.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, pub_HostName & "-" & Me.Name & "-2", Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   End If
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set rsQuery = Nothing
   Set frm010027 = Nothing
End Sub

Private Function ChkLogFile(pLogFile As String, Optional pResult As Integer) As Boolean
   Dim strText As String
   
   strText = ReadTextFile(pLogFile, "UTF-8")
   If InStr(strText, "查詢結果無可簽收的") > 0 Then
      ChkLogFile = True
      pResult = 0
   ElseIf InStr(strText, "未約定電子送達") > 0 Then
      pResult = 1
   Else
      pResult = -1
   End If
End Function

Private Function ReadTextFile(pFileName As String, Optional pCharset As String = "big5") As String
   Dim adoStream As ADODB.Stream
   Dim var_String As Variant
   
   Set adoStream = New ADODB.Stream
   'adoStream.Charset = "UTF-8"
   adoStream.Charset = pCharset
   adoStream.Open
   adoStream.LoadFromFile pFileName
   ReadTextFile = adoStream.ReadText
   adoStream.Close
   Set adoStream = Nothing
End Function

Private Function GetStr(ByVal pContent As String) As String
   '去除前後空白
   pContent = Trim(pContent)
   
   '最後可能會多個逗號
   If Right(pContent, 1) = "," Then
      pContent = Left(pContent, Len(pContent) - 1)
   End If
   
   '去除前後的雙引號
   Do While Left(pContent, 1) = """"
      pContent = Mid(pContent, 2)
   Loop
   Do While Right(pContent, 1) = """"
      pContent = Left(pContent, Len(pContent) - 1)
   Loop
   
   GetStr = pContent
   
End Function

Private Function CopyFile(pFromFolder As String, pToFolder As String, ByVal pFileList As String) As Boolean
   Dim arrFiles() As String
   Dim idx As Integer
   Dim stFromPath As String
   Dim stToPath As String
   
On Error GoTo ErrHnd

   'Modified by Morgan 2017/5/25 檔名會有;號
   'arrFiles = Split(pFileList, ";")
   'Modified by Morgan 2023/3/22 +LCase(檔名會有大寫)
   pFileList = Replace(LCase(pFileList), ".pdf;", ".pdf" & Chr(10))
   arrFiles = Split(pFileList, Chr(10))
   'end 2017/5/25
   For idx = LBound(arrFiles) To UBound(arrFiles)
      If arrFiles(idx) <> "" Then
         stFromPath = pFromFolder & "\" & arrFiles(idx)
         stToPath = pToFolder & "\" & arrFiles(idx)
         If oFileSys.FileExists(stFromPath) = True Then
            If PUB_ChkDir(pToFolder) = True Then
               Set oFile = oFileSys.GetFile(stFromPath)
               oFile.Copy stToPath, True
            Else
               sbAddList "目的資料夾不存在(" & pToFolder & ")"
               Exit Function
            End If
         Else
            sbAddList "來源檔案不存在(" & stFromPath & ")"
            Exit Function
         End If
      End If
   Next
   CopyFile = True
   Exit Function
   
ErrHnd:
   sbAddList "" & Err.Number & "," & Err.Description
   
End Function
'Modified by Morgan 2023/1/12 +案由:pCaseSubject
Private Function JoinPdf(ByVal pFromFiles As String, ByVal pToFileName As String, ByVal pFromPath As String, ByVal pAppNo As String, ByVal pCaseType As String, Optional ByVal pToPath As String = ".", Optional ByVal pCaseSubject As String, Optional ByVal pEDocNo As String) As Boolean
   Dim bolJoin As Boolean
   Dim arrFiles() As String
   Dim stTempName As String, idx As Integer
   Dim stNewFiles As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim bolRetry As Boolean
   Dim bolFeeForm As Boolean
   Dim stSpliter As String
   Dim ii As Integer
   Dim bolCert As Boolean, iCert As Integer, bolNoJoin As Boolean 'Added by Morgan 2023/1/12
   
On Error GoTo ErrHnd
   
   If PUB_CheckIsRunning("pdftk.exe") = True Then
      sbAddList "合併程式目前已在執行中，請先結束後才可繼續作業"
      Exit Function
   End If
      
   'Added by Morgan 2023/1/12
   bolCert = False
   iCert = 0
   'Modified by Morgan 2023/2/13 +准予讓與專利權
   If (pCaseType = "P" And (pCaseSubject = "准予讓與專利權" Or InStr(pCaseSubject, "專利證書") > 0)) _
      Or (pCaseType = "T" And InStr(pCaseSubject, "註冊") > 0) Then
      bolCert = True
   End If
   'end 2023/1/12
   
   '切換至來源目錄
   If pFromPath <> "." Then ChDir pFromPath
   
   '中文檔名無法合併要將檔案依順序重新命名為 pToFileName.001, pToFileName.002, pToFileName.003..
   'Modified by Morgan 2017/5/25 申請號會有";"號,改替換成 chr(10)
   'arrFiles = Split(pFromFiles, ";")
   stSpliter = Chr(10)
   'Modified by Morgan 2023/2/13 電子專利證書的副檔名是大寫
   'pFromFiles = Replace(pFromFiles, ".pdf;", ".pdf" & stSpliter)
   pFromFiles = Replace(LCase(pFromFiles), ".pdf;", ".pdf" & stSpliter)
   arrFiles = Split(pFromFiles, stSpliter)
   'end 2017/5/25
   stNewFiles = ""
   For idx = LBound(arrFiles) To UBound(arrFiles)
      bolNoJoin = False 'Added by Morgan 2023/1/12
      '專利申請書(檔名為SER002.pdf結尾)P案不印也不合併(不存檔),但要複製到 \\Pat1\Fee_Form\本所案號+"."+原檔名以便客戶調用
      'Modified by Morgan 2020/9/9 繳費單都不要合併,若找不到案號時放申請號,收文後再人工改檔名 Ex:FCT-046189
      bolFeeForm = False
      If pCaseType = "P" Then
         'Modified by Morgan 2020/8/11 +_IP110_1.pdf(檔案名稱有變)
         'Modified by Morgan 2022/12/30 +_IP115_,_IP116_ 領證申請書、延緩公告申請書並改模糊比對
         'If UCase(Right(arrFiles(idx), 10)) = UCase("SER002.pdf") Or UCase(Right(arrFiles(idx), 12)) = UCase("_IP110_1.pdf") Then
         'Modified by Morgan 2023/6/2  +申請書.pdf
         If UCase(Right(arrFiles(idx), 10)) = UCase("SER002.pdf") Or UCase(Right(arrFiles(idx), 12)) = UCase("_IP110_1.pdf") Or InStr(UCase(arrFiles(idx)), "_IP115_") > 0 Or InStr(UCase(arrFiles(idx)), "_IP116_") > 0 Or InStr(LCase(arrFiles(idx)), "申請書.pdf") > 0 Then
         'end 2022/12/30
            bolFeeForm = True
            'Modified by Morgan 2017/5/26 專利繳費單都不要列印(不合併)
            'strExc(0) = "select pa01||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) from patent where pa11='" & pAppNo & "' and pa01='P'"
            strExc(0) = "select pa01||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) from patent where pa11='" & pAppNo & "'"
            'end 2017/5/26
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               stTempName = cPFeeForm & "\" & RsTemp(0) & "." & arrFiles(idx)
            Else
               stTempName = cPFeeForm & "\" & pAppNo & "." & arrFiles(idx)
               PUB_SendMail strUserNum, "92012", "", "申請號 " & pAppNo & " 的繳費單無對應本所案號，請於收文後修正檔名", "繳費單目前檔名：" & stTempName
            End If
            'NAS用FileSystemObject檢查結果不穩定,用PUB_ChkDir檢查
            If m_bPFeeFormFolderCheck = False Then
               If PUB_ChkDir(cPFeeForm) = False Then
                  sbAddList "無法讀取申請書資料夾(" & cPFeeForm & ")，作業取消"
                  GoTo ErrHnd
               Else
                  m_bPFeeFormFolderCheck = True
               End If
            End If
         End If
      
      'Added by Morgan 2018/11/30
      '商標也不印繳費單,要複製到 \\Tm31\xfer\Fee_Form\本所案號+"."+原檔名以便承辦調用
      ElseIf pCaseType = "T" Then
         'Modified by Morgan 2020/8/11 +_IT002.pdf(檔案名稱有變)
         'Modified by Morgan 2021/11/15 檔名又變成_IT002_000.pdf，改用模糊比對
         'If UCase(Right(arrFiles(idx), 10)) = UCase("regfee.pdf") Or UCase(Right(arrFiles(idx), 10)) = UCase("_IT002.pdf") Then
         'Modified by Morgan 2023/6/2 發生檔名為 -regfee.pdf.pdf 狀況
         'If UCase(Right(arrFiles(idx), 10)) = UCase("regfee.pdf") Or InStr(UCase(arrFiles(idx)), "_IT002") > 0 Then
         If InStr(LCase(arrFiles(idx)), "regfee.pdf") > 0 Or InStr(UCase(arrFiles(idx)), "_IT002") > 0 Then
         'end 2023/6/2
            bolFeeForm = True
            'Modified by Morgan 2018/12/3 有做T案 , FCT維持原作業 --陳金蓮
            'Modified by Morgan 2018/12/17 FCT又改同T --陳金蓮
            strExc(0) = "select tm01||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04) from trademark where tm12='" & pAppNo & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               stTempName = strTFeeForm & "\" & RsTemp(0) & "." & arrFiles(idx)
            Else
               stTempName = strTFeeForm & "\" & pAppNo & "." & arrFiles(idx)
               PUB_SendMail strUserNum, "92012", "", "申請號 " & pAppNo & " 的繳費單無對應本所案號，請於收文後修正檔名", "繳費單目前檔名：" & stTempName
            End If
            'NAS用FileSystemObject檢查結果不穩定,用PUB_ChkDir檢查
            If m_bTFeeFormFolderCheck = False Then
               If PUB_ChkDir(strTFeeForm) = False Then
                  sbAddList "無法讀取繳費單資料夾(" & strTFeeForm & ")，作業取消"
                  GoTo ErrHnd
               Else
                  m_bTFeeFormFolderCheck = True
               End If
            End If
         End If
      'end 2018/11/30
      End If

      If bolFeeForm = False Then
         stTempName = pToFileName & "." & Format(idx + 1, "000") '序號需索引一致後面更名才會對到
      End If
            
      If oFileSys.FileExists(arrFiles(idx)) = True Then         '
         '複製申請書到指定位置
         If bolFeeForm = True Then
            'Modified by Morgan 2019/7/4 台灣商標案註冊費繳費單只存第一頁--陳金蓮,林桂英
            'Set oFile = oFileSys.GetFile(arrFiles(idx))
            If pCaseType = "T" Then
               If ShellProgram(pub_PdftkEXE, arrFiles(idx) & " cat 1 output .\" & arrFiles(idx) & ".P1") = True Then
                  'Added by Morgan 2019/7/22 加3次檢查(108/7/22發生找不到檔案53的錯誤,可能是分割程式結束但檔案尚未完成的時間差造成
                  For ii = 1 To 3
                     If oFileSys.FileExists(arrFiles(idx) & ".P1") = True Then
                        Exit For
                     Else
                        Sleep 1000
                     End If
                  Next
                  'end 2019/7/22
                  Set oFile = oFileSys.GetFile(arrFiles(idx) & ".P1")
               Else
                  Set oFile = oFileSys.GetFile(arrFiles(idx))
               End If
            Else
               Set oFile = oFileSys.GetFile(arrFiles(idx))
            End If
            'end 2019/7/4
            oFile.Copy stTempName, True
            
         Else
            Set oFile = oFileSys.GetFile(arrFiles(idx))
            'Added by Morgan 2023/1/16
            'Modified by Morgan 2023/2/1
            'If bolCert And (InStr(arrFiles(idx), "_IP") > 0 Or InStr(arrFiles(idx), "_IT") > 0) Then
            'Modified by Morgan 2023/2/13 電子專利證書目前檔名為申請號_專利號(無公文號)
            'Modified by Morgan 2023/6/13
            If bolCert And InStr(arrFiles(idx), "下載說明") = 0 And InStr(UCase(arrFiles(idx)), UCase("-issueAtt")) = 0 And (InStr(UCase(arrFiles(idx)), UCase("-certAtt")) > 0 Or InStr(arrFiles(idx), "電子") > 0 Or InStr(arrFiles(idx), "_IP003") > 0 Or InStr(arrFiles(idx), pEDocNo) = 0) Then
               stTempName = Replace(pToFileName, ".pdf", ".CERT" & IIf(iCert > 0, "." & iCert, "") & ".pdf")
               iCert = iCert + 1
               'oFile.Copy stTempName, True
               bolNoJoin = True
            End If
            'end 2023/1/16
            oFile.Name = stTempName
         End If
      '再檢查是否存在已更名檔案(前次失敗)
      ElseIf oFileSys.FileExists(stTempName) = False Then
         sbAddList "找不到pdf附件(" & arrFiles(idx) & ")，作業取消"
         GoTo ErrHnd
      End If
      
      If bolFeeForm = False And bolNoJoin = False Then
         'Modified by Morgan 2017/5/25
         'stNewFiles = IIf(stNewFiles <> "", stNewFiles & ";", "") & stTempName
         stNewFiles = IIf(stNewFiles <> "", stNewFiles & stSpliter, "") & stTempName
      End If
   Next
   
   '1個檔案用更名
   'Modified by Morgan 2017/5/25
   'If InStr(stNewFiles, ";") = 0 Then
   If InStr(stNewFiles, stSpliter) = 0 Then
   'end 2017/5/25
      Set oFile = oFileSys.GetFile(stNewFiles)
      oFile.Name = pToFileName
   '合併
   Else
   
      bolJoin = False
      stNewFiles = Replace(stNewFiles, stSpliter, " ")
      '刪舊檔
      If Dir(".\" & pToFileName) <> "" Then Kill ".\" & pToFileName
      
      '特殊路徑無法存檔故先就地合併
      If ShellProgram(pub_PdftkEXE, stNewFiles & " cat output .\" & pToFileName) = True Then
         '檢查合併檔是否可開啟
         If ChkPdfOK(".\" & pToFileName) = False Then
            Kill ".\" & pToFileName
         Else
            bolJoin = True
         End If
      End If
      
      If bolJoin = False Then
         sbAddList "PDF檔合併失敗," & "來源檔：" & pFromFiles & vbCrLf & "目的檔：" & pToFileName & vbCrLf & vbCrLf & "（若系統確實無法合併時，請自行合併來源檔為目的檔後存放於來源目錄！）"
         '檔名還原
         For idx = LBound(arrFiles) To UBound(arrFiles)
            stTempName = pToFileName & "." & Format(idx + 1, "000")
            If oFileSys.FileExists(stTempName) = True Then
               Set oFile = oFileSys.GetFile(stTempName)
               bolRetry = False
               oFile.Name = arrFiles(idx)
            End If
         Next
         GoTo ExitPoint
      End If
   End If
   
   '若目的地路徑不同
   If pToPath <> "." Then
      '搬到目的地
      If oFileSys.FileExists(pToPath & "\" & pToFileName) = True Then oFileSys.DeleteFile pToPath & "\" & pToFileName
      oFileSys.MoveFile ".\" & pToFileName, pToPath & "\" & pToFileName
   End If
      
   JoinPdf = True
   
ErrHnd:
   If Err.Number <> 0 Then
      '檔案可能不會馬上釋放,會無法更名
      If Err.Number = 70 And bolRetry = False Then
         Sleep 1000
         bolRetry = True
         Resume
      Else
         sbAddList "" & Err.Number & "," & Err.Description
      End If
   End If
   
ExitPoint:
   '目錄切回
   If InStr(App.path, "\\") = 1 Then
      ChDir "C:\"
   Else
      ChDir App.path
   End If
End Function


'檢查合併檔是否正確
Private Function ChkPdfOK(pFileName As String) As Boolean
   Dim stTmpFile As String
   Dim iTimes As Integer
   Dim strCmd As String
   
On Error GoTo ErrHnd

   stTmpFile = ".\$$Check.pdf"
   If Dir(stTmpFile) <> "" Then Kill stTmpFile
   
   strCmd = pub_PdftkEXE & " " & pFileName & " cat output " & stTmpFile
   Shell strCmd
   For iTimes = 1 To 10
      If PUB_CheckIsRunning(pub_PdftkName) = True Then
         Sleep 1000
      Else
         Exit For
      End If
   Next
   If iTimes > 10 Then Exit Function
   ChkPdfOK = True
   
ErrHnd:

End Function

'通知各單位程序人員
Private Function fnEMail() As Boolean
   Dim stMailTo As String, stMailCC As String, stMailBCC As String, stSubject As String, stContent As String, iCount As Integer
   Dim stDate As String
   
   stDate = ChangeTStringToTDateString(strSrvDate(2))
   'stMailBCC = "92012"
   
   'Modified by Morgan 2017/6/19 改寄到公用信箱 tipomsg@taie.com.tw
   'strExc(0) = "select st01 from staff where st05='M1' and st04='1' order by 1"
   'intI = 1
   'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   'If intI = 1 Then
   '   stMailCC = RsTemp.GetString(, , , ";")
   'End If
   stMailCC = "tipomsg@taie.com.tw"
   'end 2017/6/19
   
   If Combo1 = "" Or Combo1 = "P" Then
      '內專(含抓不到本所案號的專利案)
      stMailTo = Pub_GetSpecMan("電子公文匯入通知P")
      'Modified by Morgan 2017/6/22
      'strExc(0) = "select count(*) from edocument" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='P'" & _
         " and not exists(select * from patent where pa11=substr(ed02,1,9) and pa01='FCP')"
      
      strExc(0) = "select count(*) from edocument" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='P' and (substr(ed27,1,1)='P' or ed27 is null)"
      'end 2017/6/22
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            stSubject = stDate & " 內專電子公文已匯入(" & RsTemp(0) & "筆)"
         Else
            stSubject = stDate & " 內專無電子公文"
         End If
         sbAddList "通知內專[" & stSubject & "] => " & stMailTo
         PUB_SendMail strUserNum, stMailTo, "", stSubject, " 如旨", , , , , , stMailCC, , , , , , stMailBCC
      End If
   End If
      
   'FCP
   If Combo1 = "" Or Combo1 = "FCP" Then
      stMailTo = Pub_GetSpecMan("電子公文匯入通知FCP")
      'Modified by Morgan 2017/6/22
      'strExc(0) = "select count(*) from edocument" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='P' " & _
         " and exists(select * from patent where pa11=substr(ed02,1,9) and pa01='FCP')"
      
      strExc(0) = "select count(*) from edocument" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='P' and substr(ed27,1,3)='FCP'"
      'end 2017/6/22
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            stSubject = stDate & " FCP電子公文已匯入(" & RsTemp(0) & "筆)"
         Else
            stSubject = stDate & " FCP無電子公文"
         End If
         sbAddList "通知FCP[" & stSubject & "] => " & stMailTo
         PUB_SendMail strUserNum, stMailTo, "", stSubject, " 如旨", , , , , , stMailCC, , , , , , stMailBCC
      End If
   End If
      
   'T(含抓不到本所案號的商標案)
   If Combo1 = "" Or Combo1 = "T" Then
      stMailTo = Pub_GetSpecMan("電子公文匯入通知T")
      'Modified by Morgan 2017/6/22
      'strExc(0) = "select count(*) from edocument,edoccodemap" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='T' and em01(+)=ed10 and em02(+)=ed04" & _
         " and ( em07='Y' or not (exists(select * from trademark where tm12=ed02 and tm01='FCT' and tm28='1')" & _
         " or exists(select * from trademark where tm15=ed16 and tm01='FCT' and tm28='1')" & _
         " or exists(select * from caseprogress,trademark where cp30=ed02 and cp01='FCT' and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm28='1'))" & _
         ")"
         
      strExc(0) = "select count(*) from edocument,edoccodemap" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='T' and em01(+)=ed10 and em02(+)=ed04" & _
         " and (em07='Y' or ed27 is null or  not exists(select * from trademark where tm01=substr(ed27,1,length(ed27)-9)" & _
         " and tm02=substr(ed27,length(ed27)-8,6)" & _
         " and tm03=substr(ed27,length(ed27)-2,1)" & _
         " and tm04=substr(ed27,length(ed27)-1,2)" & _
         " and tm01='FCT' and tm28='1'))"
      'end 2017/6/22
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            stSubject = stDate & " 內商電子公文已匯入(" & RsTemp(0) & "筆)"
         Else
            stSubject = stDate & " 內商無電子公文"
         End If
         sbAddList "通知內商[" & stSubject & "] => " & stMailTo
         PUB_SendMail strUserNum, stMailTo, "", stSubject, " 如旨", , , , , , stMailCC, , , , , , stMailBCC
      End If
   End If
   
   'FCT
   If Combo1 = "" Or Combo1 = "FCT" Then
      stMailTo = Pub_GetSpecMan("電子公文匯入通知FCT")
      'Modified by Morgan 2017/6/22
      'strExc(0) = "select count(*) from edocument,edoccodemap" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='T' and em01(+)=ed10 and em02(+)=ed04 and em07 is null" & _
         " and (exists(select * from trademark where tm12=ed02 and tm01='FCT' and tm28='1')" & _
         " or exists(select * from trademark where tm15=ed16 and tm01='FCT' and tm28='1')" & _
         " or exists(select * from caseprogress,trademark where cp30=ed02 and cp01='FCT' and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm28='1')" & _
         ")"
      
      strExc(0) = "select count(*) from edocument,edoccodemap" & _
         " Where trunc(ed05)=trunc(sysdate) and ed10='T' and em01(+)=ed10 and em02(+)=ed04 and em07 is null and ed27 is not null" & _
         " and exists(select * from trademark where tm01=substr(ed27,1,length(ed27)-9)" & _
         " and tm02=substr(ed27,length(ed27)-8,6)" & _
         " and tm03=substr(ed27,length(ed27)-2,1)" & _
         " and tm04=substr(ed27,length(ed27)-1,2)" & _
         " and tm01='FCT' and tm28='1')"
      'end 2017/6/22
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If RsTemp(0) > 0 Then
            stSubject = stDate & " 外商電子公文已匯入(" & RsTemp(0) & "筆)"
         Else
            stSubject = stDate & " 外商無電子公文"
         End If
         sbAddList "通知外商[" & stSubject & "] => " & stMailTo
         PUB_SendMail strUserNum, stMailTo, "", stSubject, " 如旨", , , , , , stMailCC, , , , , , stMailBCC
      End If
   End If
   
   '簽收統計
   '重新下載的CSV檔沒有收受人
   '目前智慧局的發文號都是11碼,12碼的都是系統新增的(如一文多案或分割核准)
   If Check1.Value = vbChecked Then
      'INSTR(ED31,'代理人：')
      'Modified by Morgan 2020/4/23 台一的商標案也有桂所長名字但非代理人,改判斷>INSTR(ED31,'代理人：')
      'Modified by Morgan 2024/1/3 因有些案件林總是複委任代理人，智慧局公文不一定會將其列為案件代理人；且桂所長已退休，故改只統計總數量，不再以代理人名統計。
      'strExc(0) = "SELECT '桂齊恆' 代理人,decode(ed10,'P','專利','商標') 種類,count(*) 數量 FROM EDOCUMENT" & _
         " WHERE ED05>TRUNC(SYSDATE) AND (INSTR(ED31,'桂齊恆')+INSTR(ED31,'桂齊恒'))>INSTR(ED31,'代理人：')" & _
         " and length(ed01)=11 group by ed10" & _
         " Union All" & _
         " SELECT '林景郁' 代理人,decode(ed10,'P','專利','商標') 種類,count(*) 數量 FROM EDOCUMENT" & _
         " WHERE ED05>TRUNC(SYSDATE) AND INSTR(ED31,'林景郁')>INSTR(ED31,'代理人：')" & _
         " and length(ed01)=11" & _
         " group by ed10"
      strExc(0) = " SELECT decode(ed10,'P','專利','商標') 種類,count(*) 數量 FROM EDOCUMENT" & _
         " WHERE ED05>TRUNC(SYSDATE)  and length(ed01)=11 group by ed10"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         iCount = 0
         stContent = "本統計請與智慧局通知信檢核，不要與各部門匯入筆數比對!!!" & vbCrLf
         Do While Not RsTemp.EOF
            'stContent = stContent & vbCrLf & RsTemp("代理人") & "　" & RsTemp("數量") & "　筆[" & RsTemp("種類") & "]"
            stContent = stContent & vbCrLf & RsTemp("種類") & "　" & RsTemp("數量")
            iCount = iCount + RsTemp("數量")
            RsTemp.MoveNext
         Loop
         
         'Added by Morgan 2019/7/1
         If m_bolContinue = True Then
            stContent = stContent & vbCrLf & vbCrLf & "印表機待列印文件：" & m_iPCount + m_iLCount & "(清單：" & m_iLCount & ",公文：" & m_iPCount & ")"
         End If
         'end 2019/7/1
         
         stSubject = stDate & " 應被通知有 " & iCount & " 筆電子公文待簽收"
         
         sbAddList "通知櫃台[" & stSubject & "] => " & stMailCC
         PUB_SendMail strUserNum, stMailCC, "", stSubject, stContent, , , , , , , , , , , , stMailBCC
      End If
   End If
      
   fnEMail = True
End Function

'P案調卷清單(含抓不到本所案號的專利案)
Private Sub ReportP(pDate As String, Optional pPrintAtt As Boolean = False)
   Dim stDate As String
   
   If pDate = "" Then
      stDate = strSrvDate(1)
   Else
      stDate = DBDATE(pDate)
   End If
   
   'Modified by Morgan 2017/6/22
   'strExc(0) = "select decode(pa01,'','非本所案件',pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)) 本所案號" & _
      ",ed02 申請案號,ed04||decode(ed28,'副本','-'||ed28) 案由, sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期" & _
      ",nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,ed01,cpp02,ed20,ed11" & _
      " from edocument,(select ed01 X1,ed23 X2,pa01,pa02,pa03,pa04,pa75,pa26 from edocument e,patent" & _
      " where trunc(ed05)=to_date(" & stDate & ",'YYYYMMDD') and ed10='P' and instr(ed02,'N')=0 and pa11(+)=ed02 and pa01 in ('P','FCP')" & _
      " union select ed01 X1,ed23 X2,pa01,pa02,pa03,pa04,pa75,pa26 from edocument e,caseprogress,patent" & _
      " where trunc(ed05)=to_date(" & stDate & ",'YYYYMMDD') and ed10='P' and instr(ed02,'N')>0 and cp36(+)=ed02 and cp01 in ('P','FCP')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " union select ed01 X1,ed23 X2,pa01,pa02,pa03,pa04,pa75,pa26 from edocument e,patent" & _
      " where trunc(ed05)=to_date(" & stDate & ",'YYYYMMDD') and ed10='P' and instr(ed02,'N')>0 and pa11(+)=substr(ed02,1,9) and pa01 in ('P','FCP')" & _
      " and not exists(select * from caseprogress where cp36=ed02 and cp01 in ('P','FCP'))" & _
      ") X,casepaperpdf where trunc(ed05)=to_date(" & stDate & ",'YYYYMMDD') and ed10='P' and X1(+)=ed01 and X2(+)=ed23 and (pa01='P' or pa01 is null) and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " order by pa01,pa02,pa03,pa04,ed01"
   
   strExc(0) = "select decode(pa01,'','非本所案件',pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)) 本所案號" & _
      ",ed02 申請案號,ed04||decode(ed28,'副本','-'||ed28) 案由, sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期" & _
      ",nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,ed01,cpp02,ed20,ed11" & _
      " from edocument,patent,casepaperpdf" & _
      " where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='P'" & _
      " and pa01(+)=substr(ed27,1,length(ed27)-9)" & _
      " and pa02(+)=substr(ed27,length(ed27)-8,6)" & _
      " and pa03(+)=substr(ed27,length(ed27)-2,1)" & _
      " and pa04(+)=substr(ed27,length(ed27)-1,2)" & _
      " and (pa01='P' or pa01 is null) and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " order by pa01,pa02,pa03,pa04,ed01,ed23"
   'end 2017/6/22
   
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      sbAddList "列印內專清單"
      DoPrint rsQuery, pDate, "P"
      '印附件
      If pPrintAtt Then
         sbAddList "列印內專附件"
         sbPrintAtt rsQuery
      End If
   Else
      sbAddList "無可列印資料"
   End If
End Sub

'FCP案調卷清單(剔除不調卷案由)
Public Sub ReportFCP(pDate As String, Optional pPrintAtt As Boolean = False, Optional pUserNo As String)
   Dim stDate As String, stCon As String
   
   'Removed by Morgan 2020/2/4 不調卷公文改各區自行輸入
   'Dim stXGirl As String
   'stXGirl = Pub_GetSpecMan("FCP來函整理人員")
   'end 2020/2/4
   
   If pDate = "" Then
      stDate = strSrvDate(1)
   Else
      stDate = TransDate(pDate, 2)
   End If
   
   If pUserNo <> "" Then stCon = " and na16='" & pUserNo & "'" 'Added by Morgan 2020/4/9
   
   'Modified by Morgan 2017/6/22
   'strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) 本所案號" & _
      ",st02 程序分區,ed02 申請案號,ed04||decode(ed28,'副本','-'||ed28) 案由,nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,em07 不調卷" & _
      ",ed01,cpp02,ed20,ed11,pa75,pa26,decode(em07,'N','" & stXGirl & "',na16) IUser,pa01||pa02||pa03||pa04 CaseNo,pa57" & _
      " from (select ed01 X1,ed23 X2,pa01,pa02,pa03,pa04,pa75,pa26,pa57 from edocument e,patent" & _
      " where trunc(ed05)=to_date(" & stDate & ",'YYYYMMDD') and ed10='P' and instr(ed02,'N')=0 and pa11(+)=ed02 and pa01='FCP'" & _
      " union select ed01 X1,ed23 X2,pa01,pa02,pa03,pa04,pa75,pa26,pa57 from edocument e,caseprogress,patent" & _
      " where trunc(ed05)=to_date(" & stDate & ",'YYYYMMDD') and ed10='P' and instr(ed02,'N')>0 and cp36(+)=ed02 and cp01='FCP'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " union select ed01 X1,ed23 X2,pa01,pa02,pa03,pa04,pa75,pa26,pa57 from edocument e,patent" & _
      " where trunc(ed05)=to_date(" & stDate & ",'YYYYMMDD') and ed10='P' and instr(ed02,'N')>0 and pa11(+)=substr(ed02,1,9) and pa01='FCP'" & _
      " and not exists(select * from caseprogress where cp36=ed02 and cp01='FCP')" & _
      ") X,edocument,edoccodemap,casepaperpdf,fagent,nation,staff where ed01(+)=X1 and ed23(+)=X2 and em01(+)=ed10 and em02(+)=ed04 and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and na01(+)=fa10 and st01(+)=na16" & _
      " order by pa01,pa02,pa03,pa04,ed01"
   'Modified by Morgan 2020/2/4 不調卷公文改各區自行輸入 decode(em07,'N','" & stXGirl & "',na16) IUser-->na16 IUser
   'Modified by Morgan 2020/6/17 +通知實審
   strExc(0) = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) 本所案號" & _
      ",st02 程序分區,ed02 申請案號,ed04||decode(ed28,'副本','-'||ed28) 案由,nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,em07 不調卷" & _
      ",ed01,cpp02,ed20,ed11,pa75,pa26,na16 IUser,pa01||pa02||pa03||pa04 CaseNo,pa57,decode(em03||em04,em03,em03) cp10" & _
      " from edocument,patent,edoccodemap,casepaperpdf,fagent,nation,staff" & _
      " where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='P'" & _
      " and pa01(+)=substr(ed27,1,length(ed27)-9)" & _
      " and pa02(+)=substr(ed27,length(ed27)-8,6)" & _
      " and pa03(+)=substr(ed27,length(ed27)-2,1)" & _
      " and pa04(+)=substr(ed27,length(ed27)-1,2)" & _
      " and pa01='FCP' and em01(+)=ed10 and em02(+)=ed04 and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " and fa01(+)=substr(pa75,1,8) and fa02(+)=substr(pa75,9) and na01(+)=fa10 and st01(+)=na16" & stCon & _
      " order by pa01,pa02,pa03,pa04,ed01,ed23"
   'end 2017/6/22
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pUserNo <> "" Then
         rsQuery.Sort = "IUser asc,不調卷 desc,CaseNo asc"
         DoPrint rsQuery, pDate, "FCP", , True
      Else
      
'Modified by Morgan 2021/11/30 改不印公文及清單，只印櫃檯核對用的清單--Sharon
'         '調卷清單
'         sbAddList "列印FCP清單(要調卷)"
'         DoPrint rsQuery, pDate, "FCP"
'
'         '印附件
'         If pPrintAtt Then
'            sbAddList "列印FCP附件(要調卷)"
'            sbPrintAtt rsQuery, "FCP"
'         End If
'
'         rsQuery.Sort = "IUser asc,不調卷 desc,CaseNo asc" 'Added by Morgan 2020/2/4 以管制人排序，是否調卷
'
'         '不調卷清單
'         sbAddList "列印FCP清單(不調卷)"
'         DoPrint rsQuery, pDate, "FCP", True
'
'         '印附件
'         If pPrintAtt Then
'            sbAddList "列印FCP附件(不調卷)"
'            sbPrintAtt rsQuery, "FCP", True
'         End If
'
'         '各區清單
'         'rsQuery.Sort = "IUser asc,CaseNo asc" 'Removed by Morgan 2020/2/4 移到上面
'         sbAddList "列印FCP清單(各區)"
'         DoPrint rsQuery, pDate, "FCP", , True
         
         sbAddList "列印FCP清單"
         DoPrint rsQuery, pDate, "FCP", , , True
'end 2021/11/30

      End If
      
   Else
      If pUserNo <> "" Then
         MsgBox "無公文！", vbExclamation
      Else
         sbAddList "FCP無公文"
      End If
   End If
End Sub

'T案調卷清單(含抓不到本所案號的商標案)
Private Sub ReportT(pDate As String, Optional pPrintAtt As Boolean = False)
   Dim stDate As String
   
   If pDate = "" Then
      stDate = strSrvDate(1)
   Else
      stDate = DBDATE(pDate)
   End If
   
   'Modified by Morgan 2017/6/22
   'strExc(0) = "select decode(tm01,'','非本所案件',tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04)) 本所案號" & _
      ",nvl(nvl(tm15,tm12),nvl(ed16,ed02)) ""申請/註冊號"",ed04||decode(ed28,'副本','-'||ed28) 案由, sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期" & _
      ",nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,ed01,cpp02,ed20,ed11" & _
      " from edocument,EDocCodeMap,(select ed01 X1,ed23 X2,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from edocument,trademark" & _
      " Where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T' and tm12(+)=ed02 and tm01 is not null and (tm15 is null or tm15=tm16)" & _
      " union select ed01 X1,ed23 X2,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from edocument,trademark" & _
      " Where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T' and tm15(+)=ed16 and tm01 is not null" & _
      " union select ed01 X1,ed23 X2,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from edocument,caseprogress,trademark" & _
      " Where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T' and cp30(+)=ed02" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
      ") X,casepaperpdf" & _
      " where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T' and em01(+)=ed10 and em02(+)=ed04" & _
      " and X1(+)=ed01 and X2(+)=ed23 and (tm01='T' or tm28<>'1' or em07='Y' or tm01 is null) and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " order by tm01,tm02,tm03,tm04,ed01"
      
   strExc(0) = "select decode(tm01,'','非本所案件',tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04)) 本所案號" & _
      ",nvl(nvl(tm15,tm12),nvl(ed16,ed02)) ""申請/註冊號"",ed04||decode(ed28,'副本','-'||ed28) 案由, sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期" & _
      ",nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,ed01,cpp02,ed20,ed11" & _
      " from edocument,trademark,EDocCodeMap,casepaperpdf" & _
      " where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T'" & _
      " and tm01(+)=substr(ed27,1,length(ed27)-9)" & _
      " and tm02(+)=substr(ed27,length(ed27)-8,6)" & _
      " and tm03(+)=substr(ed27,length(ed27)-2,1)" & _
      " and tm04(+)=substr(ed27,length(ed27)-1,2)" & _
      " and em01(+)=ed10 and em02(+)=ed04 and (tm01='T' or tm28<>'1' or em07='Y' or tm01 is null) and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " order by tm01,tm02,tm03,tm04,ed01,ed23"
   'end 2017/6/22
   
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      sbAddList "列印內商清單"
      DoPrint rsQuery, pDate, "T"
      '印附件
      If pPrintAtt Then
         sbAddList "列印內商附件"
         sbPrintAtt rsQuery
      End If
      PrintDivList pDate, "T"
   Else
      sbAddList "內商無公文"
   End If
End Sub

'FCT案調卷清單
Private Sub ReportFCT(pDate As String, Optional pPrintAtt As Boolean = False)
   Dim stDate As String
   
   If pDate = "" Then
      stDate = strSrvDate(1)
   Else
      stDate = DBDATE(pDate)
   End If
   
   'Modified by Morgan 2017/6/22
   'strExc(0) = "select tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04) 本所案號" & _
      ",nvl(tm15,tm12) ""申請/註冊號"",ed04||decode(ed28,'副本','-'||ed28) 案由, sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期" & _
      ",nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,ed01,cpp02,ed20,ed11" & _
      " from (select ed01 X1,ed23 X2,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from edocument,trademark" & _
      " Where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T' and tm12(+)=ed02 and tm01 is not null" & _
      " union select ed01 X1,ed23 X2,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from edocument,trademark" & _
      " Where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T' and tm15(+)=ed16 and tm01 is not null" & _
      " union select ed01 X1,ed23 X2,tm01,tm02,tm03,tm04,tm12,tm15,tm28 from edocument,caseprogress,trademark" & _
      " Where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T' and cp30(+)=ed02" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
      ") X,edocument,EDocCodeMap,casepaperpdf" & _
      " where ed01(+)=X1 and ed23(+)=X2 and em01(+)=ed10 and em02(+)=ed04 and tm01='FCT' and tm28='1' and em07 is null and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " order by tm01,tm02,tm03,tm04,ed01"
   
   strExc(0) = "select tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04) 本所案號" & _
      ",nvl(tm15,tm12) ""申請/註冊號"",ed04||decode(ed28,'副本','-'||ed28) 案由, sqldatet(to_char(ed05,'yyyymmdd')) 簽收日期" & _
      ",nvl(ed19,sqldatet(ed18)) 處理期限,sqldatet(ed08) 發文日期,ed01,cpp02,ed20,ed11" & _
      " from edocument,trademark,EDocCodeMap,casepaperpdf" & _
      " where trunc(ed05) = to_date(" & stDate & ",'YYYYMMDD') and ed10='T'" & _
      " and tm01(+)=substr(ed27,1,length(ed27)-9)" & _
      " and tm02(+)=substr(ed27,length(ed27)-8,6)" & _
      " and tm03(+)=substr(ed27,length(ed27)-2,1)" & _
      " and tm04(+)=substr(ed27,length(ed27)-1,2)" & _
      " and tm01='FCT' and tm28='1' and em01(+)=ed10 and em02(+)=ed04 and em07 is null and cpp01(+)=ed01 and cpp02(+)='$'||ed01||'.pdf'" & _
      " order by tm01,tm02,tm03,tm04,ed01,ed23"
   'end 2017/6/22
   
   intI = 1
   Set rsQuery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      sbAddList "列印外商清單"
      DoPrint rsQuery, pDate, "FCT"
      '印附件
      If pPrintAtt Then
         sbAddList "列印外商附件"
         sbPrintAtt rsQuery
      End If
      PrintDivList pDate, "FCT"
   Else
      sbAddList "外商無公文"
   End If
End Sub
'列印分割子案清單
Private Sub PrintDivList(pDate As String, pSys As String)
   Dim stSQL As String, intQ As Integer, strCaseNo As String, iCases As Integer
   Dim strTemp(2) As String
   Dim rsQuery2 As ADODB.Recordset
   
   stSQL = "select tm01||'-'||tm02||decode(tm03||tm04,'000','','-'||tm03||'-'||tm04) 母案" & _
      ",dc01||'-'||dc02||decode(dc03||dc04,'000','','-'||dc03||'-'||dc04) 子案" & _
      " from edocument,trademark,DivisionCase" & _
      " where trunc(ed05) = to_date(" & DBDATE(pDate) & ",'YYYYMMDD') and ed10='T' and ed27 is not null" & _
      " and substr(ed27,1,length(ed27)-9)='" & pSys & "'" & _
      " and tm01(+)=substr(ed27,1,length(ed27)-9)" & _
      " and tm02(+)=substr(ed27,length(ed27)-8,6)" & _
      " and tm03(+)=substr(ed27,length(ed27)-2,1)" & _
      " and tm04(+)=substr(ed27,length(ed27)-1,2) and tm28='1'" & _
      " and exists(select * From caseprogress where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp10='308' and cp27>0 and cp57 is null)" & _
      " and dc05=tm01 and dc06=tm02 and dc07=tm03 and dc08=tm04 order by 1,2"
   
   intQ = 1
   Set rsQuery2 = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      m_bolDivList = True
      If pSys = "T" Then
         sbAddList "列印內商分割子案清單"
      Else
         sbAddList "列印外商分割子案清單"
      End If
      
      With rsQuery2
      GetPleft2
      ReDim PColName(1 To UBound(PLeft))
      PColName(1) = "母案"
      PColName(2) = "子案"
      
      Printer.PaperSize = 9 'A4
      Printer.Orientation = 1 '直印
      lngPageHeight = Printer.ScaleHeight
      lngPageWidth = Printer.ScaleWidth
      lngLineHeight = 300
      Printer.Copies = 2 'Added by Morgan 2020/6/18
      PrintPageHeader
      PrintPageHeader1
      .MoveFirst
      iCases = 0
      Do While Not .EOF
         iCases = iCases + 1
         If .Fields("母案") <> strCaseNo Then
            strTemp(1) = .Fields("母案")
            strCaseNo = .Fields("母案")
         Else
            strTemp(1) = ""
         End If
         strTemp(2) = .Fields("子案")
         PrintDetail strTemp
         .MoveNext
      Loop
      Call PrintReportFooter(0, iCases)
      Printer.EndDoc
      sbAddList "完成 (" & iCases & "筆)", , True
      m_iLCount = m_iLCount + 1 'Added by Morgan 2019/6/18
      End With
   End If
End Sub

'列印附件
Private Sub sbPrintAtt(pRecordset As ADODB.Recordset, Optional pSys As String, Optional pNoFileCase As Boolean = False)
   Dim stAttTempFolder As String, stFile As String, iCount As Integer
   Dim bolPrint As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stPdfFile As String 'Added by Morgan 2024/1/18
   
   stAttTempFolder = App.path & "\EDocAtt"
   If Dir(stAttTempFolder, vbDirectory) = "" Then
      MkDir stAttTempFolder
   ElseIf Dir(stAttTempFolder & "\*.*") <> "" Then
      Kill stAttTempFolder & "\*.*"
   End If
   
   pRecordset.MoveFirst
   iCount = 0
   With pRecordset
   Do While Not .EOF
      bolPrint = True
      If pSys = "FCP" Then
         'Modified by Morgan 2020/2/25 不調卷有例外要調卷
         'If (pNoFileCase And IsNull(.Fields("不調卷"))) Or (pNoFileCase = False And .Fields("不調卷") = "N") Then
         '   bolPrint = False
         'End If
         '不調卷
         If pNoFileCase Then
            If IsNull(.Fields("不調卷")) Then
               bolPrint = False
            ElseIf fnFCPXCase("" & .Fields("cp10"), .Fields("pa75"), .Fields("pa26")) = True Then
               bolPrint = False
            End If
         '要調卷
         Else
            If .Fields("不調卷") = "N" Then
               If fnFCPXCase("" & .Fields("cp10"), .Fields("pa75"), .Fields("pa26")) = False Then
                  bolPrint = False
               End If
            End If
         End If
         'end 2020/2/25
      End If
      If bolPrint Then
         iCount = iCount + 1
         If .Fields("ed20") = "N" Then
            sbAddList "列印 " & .Fields("cpp02") & "...不收文列印取消" & iCount
         
         'Added by Morgan 2018/5/10 若已收文要改用收文號抓(歸卷還要判斷上傳日為簽收日)
         ElseIf .Fields("ed11") <> "C" Then
            sbAddList "發文號 " & .Fields("ed01") & "...已收文 " & iCount
            stSQL = "select cpp02,1 Srt from casepaperpdf where cpp01='" & .Fields("ed11") & "' and cpp15=0 and cpp02 like '$" & .Fields("ed01") & ".pdf.%'" & _
               " union all select cpp02,2 Srt from edocument,casepaperpdf where ed01='" & .Fields("ed01") & "' and cpp01(+)=ed11 and cpp06=ed13 and cpp15=0 and substr(cpp02,-4)='.pdf' order by 2,1"
            intQ = 1
            Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
            If intQ = 1 Then
               Do While Not rsQuery.EOF
                  If rsQuery.AbsolutePosition > 1 And rsQuery(1) = 2 Then Exit Do '優先印原始檔，沒有時才印合併檔
                  stFile = rsQuery(0)
                  If PUB_GetAttachFile_CPP(.Fields("ed11"), stFile, stAttTempFolder) = True Then
                     sbAddList "列印 " & rsQuery(0)
                     SetPdfFile stFile 'Added by Morgan 2024/1/18
                     If PrintOnePdf(m_PdfReader, " /n /t """ & stFile & """ """ & cmbPrinter.Text & """") = True Then
                        m_iPCount = m_iPCount + 1
                        sbAddList "完成 " & iCount & "/" & m_iPCount, , True
                     Else
                        sbAddList "失敗 " & iCount, , True
                     End If
                  Else
                     sbAddList "附件 " & rsQuery(0) & " 下載失敗"
                  End If
                  rsQuery.MoveNext
               Loop
            End If
         'end 2018/5/10
         Else
            stFile = "" & .Fields("cpp02")
            If stFile = "" Then
               If .Fields("ed11") <> "C" Then
                  sbAddList "發文號 " & .Fields("ed01") & "...已收文 " & iCount
               Else
                  sbAddList "發文號 " & .Fields("ed01") & " 附件下載失敗 " & iCount
               End If
            Else
               'Modified by Morgan 2023/1/16
               'stSQL = "select cpp02 from casepaperpdf where cpp01='" & .Fields("ed01") & "' and cpp02<>'" & .Fields("cpp02") & "' and  order by 1"
               stSQL = "select cpp02 from casepaperpdf where cpp01='" & .Fields("ed01") & "' and cpp02<>'" & .Fields("cpp02") & "'  and upper(cpp02) not like '$%.CERT.%' order by 1"
               intQ = 1
               Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
               If intQ = 1 Then
                  Do While Not rsQuery.EOF
                     stFile = rsQuery(0)
                     If PUB_GetAttachFile_CPP(.Fields("ed01"), stFile, stAttTempFolder) = True Then
                        sbAddList "列印 " & rsQuery(0)
                        SetPdfFile stFile 'Added by Morgan 2024/1/18
                        If PrintOnePdf(m_PdfReader, " /n /t """ & stFile & """ """ & cmbPrinter.Text & """") = True Then
                           m_iPCount = m_iPCount + 1
                           sbAddList "完成 " & iCount & "/" & m_iPCount, , True
                        Else
                           sbAddList "失敗 " & iCount, , True
                        End If
                     Else
                        sbAddList "附件 " & rsQuery(0) & " 下載失敗"
                     End If
                     rsQuery.MoveNext
                  Loop
                  
               Else
                  If PUB_GetAttachFile_CPP(.Fields("ed01"), stFile, stAttTempFolder) = True Then
                     sbAddList "列印 " & .Fields("cpp02")
                     SetPdfFile stFile 'Added by Morgan 2024/1/18
                     If PrintOnePdf(m_PdfReader, " /n /t """ & stFile & """ """ & cmbPrinter.Text & """") = True Then
                        m_iPCount = m_iPCount + 1
                        sbAddList "完成 " & iCount & "/" & m_iPCount, , True
                     Else
                        sbAddList "失敗 " & iCount, , True
                     End If
                  Else
                     sbAddList "附件 " & .Fields("cpp02") & " 下載失敗 " & iCount
                  End If
               End If
               
               'Added by Morgan 2023/1/16 證書
               stSQL = "select cpp02 from casepaperpdf where cpp01='" & .Fields("ed01") & "' and upper(cpp02) like '$%.CERT.%' order by 1"
               intQ = 1
               Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
               If intQ = 1 Then
                  Do While Not rsQuery.EOF
                     stFile = rsQuery(0)
                     If PUB_GetAttachFile_CPP(.Fields("ed01"), stFile, stAttTempFolder) = True Then
                        sbAddList "列印 " & rsQuery(0)
                        PUB_WaitUntilNoJob cmbPrinter.Text 'Added by Morgan 2023/4/6 要先檢查公文印表機已清空才印證書，否則順序可能會不正確
                        SetPdfFile stFile 'Added by Morgan 2024/1/18
                        If PrintOnePdf(m_PdfReader, " /n /t """ & stFile & """ """ & cmbPrinter2.Text & """") = True Then
                           m_iPCount = m_iPCount + 1
                           sbAddList "完成 " & iCount & "/" & m_iPCount, , True
                        Else
                           sbAddList "失敗 " & iCount, , True
                        End If
                     Else
                        sbAddList "附件 " & rsQuery(0) & " 下載失敗"
                     End If
                     rsQuery.MoveNext
                  Loop
                  PUB_WaitUntilNoJob cmbPrinter2.Text 'Added by Morgan 2023/4/6 要先檢查證書印表機已清空才繼續，否則順序可能會不正確
               End If
               'end 2023/1/16
               
            End If
         End If
      End If
      .MoveNext
   Loop
   End With
   
   Set rsQuery = Nothing
End Sub

Private Function PrintOnePdf(ByVal program_name As String, parameters As String) As Boolean

   Dim process_id As Long
   Dim process_handle As Long
    ' Start the program.
    On Error GoTo ShellError

    'Modified by Morgan 2017/5/15 路徑可能含空白,改加雙引號
    process_id = Shell("""" & program_name & """ " & parameters, vbHide)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
    End If
   
    PrintOnePdf = True
    Exit Function

ShellError:
    sbAddList Err.Number & ":" & Err.Description & "(" & program_name & ")"
End Function

'Modified by Morgan 2021/11/30 +pALLCase
Private Sub DoPrint(pRecordset As ADODB.Recordset, pRptDate As String, pSys As String, Optional pNoFileCase As Boolean = False, Optional pByIUser As Boolean = False, Optional pAllCase As Boolean = False)
   
   Dim strTemp() As String
   Dim iCol As Integer
   Dim iRecs As Integer, iCases As Integer
   Dim bPaper As Boolean
   Dim bSKip As Boolean
   Dim strCaseNo As String
   Dim iNoPaper As Integer '不調卷公文數 Added by Morgan 2020/2/4
   Dim iPaperCases As Integer '調卷數 Added by Morgan 2020/2/4
   
   m_bolDivList = False
   m_strRptDate = pRptDate
   m_strRptSys = pSys
   m_bolNoFileCase = pNoFileCase
   m_bolAllCase = pAllCase 'Added by Morgan 2021/11/30
   m_strIUser = ""
   
   Printer.PaperSize = 9 'A4
   Printer.Orientation = 1 '直印
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   
   'Modified by Morgan 2020/6/16 調卷清單多印一份檔案室要留底
   'Modified by Morgan 2021/11/30 + Or pALLCase = True
   If pSys = "P" Or pNoFileCase = True Or pByIUser = True Or pAllCase = True Then
      Printer.Copies = 1
      
   ElseIf pSys = "FCP" Then
      Printer.Copies = 3
      
   Else
      Printer.Copies = 2
   End If
   'end 2020/6/16
   
   With pRecordset
      GetPleft pSys
      ReDim PColName(1 To UBound(PLeft))
      For iCol = 1 To UBound(PLeft)
         PColName(iCol) = .Fields(iCol - 1).Name
      Next
      
      ReDim strTemp(1 To UBound(PLeft))
            
      .MoveFirst
      
      iPage = 1
      iRecs = 0
      iCases = 0
      iNoPaper = 0 'Added by Morgan 2020/2/4
      iPaperCases = 0 'Added by Morgan 2020/2/4
      If pByIUser Then
         m_strIUser = .Fields("IUser")
         sbAddList "列印 " & m_strIUser & " 清單"
      End If
      PrintPageHeader
      PrintPageHeader1
      Do While Not .EOF
      
         bSKip = False

         'Modified by Morgan 2021/11/30 + And pALLCase = False
         If pSys = "FCP" And pByIUser = False And pAllCase = False Then
            '不調卷清單
            If pNoFileCase = True Then
               If IsNull(.Fields("不調卷")) Then
                  bSKip = True
               'Added by Morgan 2020/2/21 排除 例外要調卷
               ElseIf fnFCPXCase("" & .Fields("cp10"), .Fields("pa75"), .Fields("pa26")) = True Then
                  bSKip = True
               'end 2020/2/21
               End If
               
            '要調卷清單
            Else
               If .Fields("不調卷") = "N" Then
                  'Added by Morgan 2020/2/21 增加 例外要調卷
                  If fnFCPXCase("" & .Fields("cp10"), .Fields("pa75"), .Fields("pa26")) = False Then
                     bSKip = True
                  End If
               End If
            End If
         End If
         
         If bSKip = False Then
            If pByIUser Then
               If .Fields("IUser") <> m_strIUser Then
                  'Modified by Morgan 2020/2/4 FCP各區加印 調卷、不調卷 統計
                  If pSys = "FCP" And pByIUser Then
                     Call PrintReportFooter(iRecs, iCases, iPaperCases, iNoPaper)
                  Else
                     Call PrintReportFooter(iRecs, iCases)
                  End If
                  'end 2020/2/4
                  
                  '考慮雙面列印,各區分文件列印
                  'Printer.NewPage
                  Printer.EndDoc
                  sbAddList "完成 (" & iRecs & "筆)", , True
                  m_iLCount = m_iLCount + 1 'Added by Morgan 2019/6/18
                  
                  iPage = 1
                  iRecs = 0
                  iCases = 0
                  iNoPaper = 0 'Added by Morgan 2020/2/4
                  iPaperCases = 0 'Added by Morgan 2020/2/4
                  strCaseNo = ""
                  m_strIUser = .Fields("IUser")
                  
                  sbAddList "列印 " & m_strIUser & " 清單"
                  PrintPageHeader
                  PrintPageHeader1
               End If
            End If
         
            iRecs = iRecs + 1
            For iCol = 1 To UBound(strTemp)
               strTemp(iCol) = "" & .Fields(iCol - 1)
               If pSys = "P" Then
                  If PColName(iCol) = "本所案號" Then
                     If PUB_GetEMailFlag(Replace(strTemp(iCol), "-", ""), , , bPaper) = True And bPaper = False Then
                        strTemp(iCol) = strTemp(iCol) & "＊"
                     End If
                  End If
                  
               ElseIf pSys = "FCP" Then
                  If PColName(iCol) = "本所案號" Then
                     If pByIUser And .Fields("pa57") = "Y" Then
                        strTemp(iCol) = strTemp(iCol) & "(閉)"
                     End If
                     
                     If .Fields("不調卷") = "N" Then
                        'Added by Morgan 2020/2/24
                        If fnFCPXCase("" & .Fields("cp10"), .Fields("pa75"), .Fields("pa26")) = True Then
                           strTemp(iCol) = strTemp(iCol) & "◇"
                        Else
                        'end 2020/2/24
                        
                           strTemp(iCol) = strTemp(iCol) & "◆"
                           iNoPaper = iNoPaper + 1
                           
                        End If 'Added by Morgan 2020/2/24
                     End If
                     
                  ElseIf PColName(iCol) = "處理期限" Then
                     'Modify By Sindy 2017/12/11 改用Table記錄備註
                     'Modify By Sindy 2018/7/18 pSys ==> .Fields("本所案號")
                     strTemp(iCol) = PUB_ReadIPOListMemo(.Fields("本所案號"), "" & .Fields("pa75"), "" & .Fields("pa26"), "" & .Fields("案由"), "" & .Fields("處理期限")) & strTemp(iCol)
                     '2017/12/11 END
'                     If InStr("" & .Fields("案由"), "核准") = 0 And Not IsNull(.Fields("處理期限")) Then
'                        '先正達 有期限來函
'                        'Y20656 (Lerner)+X7072201(Tessera, Inc.) & X70286(Invensas Corporation) 有期限來函
'                        'Y53942 Tessera 有期限來函
'                        'Modify By Sindy 2017/8/3 Y339400(Foley&Lardner,LLP)+申請人X48991 INTERSIL AMERICAS INC.
'                        If Not IsNull(.Fields("pa75")) And (InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left("" & .Fields("pa75"), 8)) > 0 _
'                           Or (Left("" & .Fields("pa75"), 8) = "Y2065600" And Not IsNull(.Fields("pa26")) And (Left("" & .Fields("pa26"), 8) = "X7072201" Or Left("" & .Fields("pa26"), 6) = "X70286")) _
'                           Or (Left("" & .Fields("pa75"), 8) = "Y3394000" And Not IsNull(.Fields("pa26")) And Left("" & .Fields("pa26"), 8) = "X4899100") _
'                           Or Left("" & .Fields("pa75"), 8) = "Y5394200") Then
'                           strTemp(iCol) = "*" & strTemp(iCol)
'
'                        'DOW 有期限來函
'                        ElseIf Left("" & .Fields("pa75"), 8) = "Y2245700" Then
'                           strTemp(iCol) = "**" & strTemp(iCol)
'
'                        ElseIf Not IsNull(.Fields("pa26")) And InStr("X6740200,X6740201,X6740202,X6050700,X6050701,X7074900", Left("" & .Fields("pa26"), 8)) > 0 Then
'                           strTemp(iCol) = "**" & strTemp(iCol)
'
'                        'Sandvik OA 需2日內報告
'                        ElseIf Not IsNull(.Fields("pa75")) And _
'                           (InStr("Y5285900", Left("" & .Fields("pa75"), 8)) > 0 Or InStr("Y5179901", Left("" & .Fields("pa75"), 8)) > 0) Then
'                           strTemp(iCol) = "**" & strTemp(iCol)
'
'                        'UNIUS 有期限來函
'                        ElseIf Left("" & .Fields("pa75"), 8) = "Y5150800" Then
'                           strTemp(iCol) = "***" & strTemp(iCol)
'                        End If
'
'                     ElseIf Not IsNull(.Fields("處理期限")) Then
'                        'Add By Sindy 2017/10/20
'                        '<Y47453> Shiga International Patent Office + <X55778> Nippon Soda Co., Ltd.
'                        '需2日內報告
'                        If Not IsNull(.Fields("pa75")) And _
'                           Not IsNull(.Fields("pa26")) And _
'                           (InStr("Y4745300", Left("" & .Fields("pa75"), 8)) > 0 And InStr("X5577800", Left("" & .Fields("pa26"), 8)) > 0) Then
'                           strTemp(iCol) = "**" & strTemp(iCol)
'                        '2017/10/20 END
'                        End If
'                     End If
                     
                  End If
               End If
               
               '本所案號相同不要印空白
               If PColName(iCol) = ("本所案號") Then
                  If .Fields("本所案號") = strCaseNo Then
                     strTemp(iCol) = ""
                  Else
                     iCases = iCases + 1
                     
                     'Added by Morgan 2020/2/4
                     If pSys = "FCP" Then
                        If IsNull(.Fields("不調卷")) Then
                           iPaperCases = iPaperCases + 1
                        'Added by Morgan 2020/2/24
                        ElseIf fnFCPXCase("" & .Fields("cp10"), .Fields("pa75"), .Fields("pa26")) = True Then
                           iPaperCases = iPaperCases + 1
                        'end 2020/2/24
                        End If
                     End If
                     'end 2020/2/4
                  End If
               End If
            Next
            PrintDetail strTemp
            strCaseNo = .Fields("本所案號")
         End If
         .MoveNext
      Loop
      'Modified by Morgan 2020/2/4 FCP各區加印 調卷、不調卷 統計
      If pSys = "FCP" And pByIUser Then
         Call PrintReportFooter(iRecs, iCases, iPaperCases, iNoPaper)
      Else
         Call PrintReportFooter(iRecs, iCases)
      End If
      'end 2020/2/4
      Printer.EndDoc
      sbAddList "完成 (" & iRecs & "筆)", , True
      m_iLCount = m_iLCount + 1 'Added by Morgan 2019/6/18
   End With
   
End Sub

Private Sub GetPleft(Optional pSys As String)
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(1 To 6)
   PLeft(1) = ciStartX
   If pSys = "FCP" Then
      PLeft(2) = PLeft(1) + Printer.TextWidth(String(8, "　")) + ciColGap
      PLeft(3) = PLeft(2) + Printer.TextWidth(String(4, "　")) + ciColGap
      PLeft(4) = PLeft(3) + Printer.TextWidth(String(6, "　")) + ciColGap
      PLeft(5) = PLeft(4) + Printer.TextWidth(String(12, "　")) + ciColGap
      PLeft(6) = PLeft(5) + Printer.TextWidth(String(6, "　")) + ciColGap
   Else
      PLeft(2) = PLeft(1) + Printer.TextWidth(String(7, "　")) + ciColGap
      PLeft(3) = PLeft(2) + Printer.TextWidth(String(6, "　")) + ciColGap
      PLeft(4) = PLeft(3) + Printer.TextWidth(String(12, "　")) + ciColGap
      PLeft(5) = PLeft(4) + Printer.TextWidth(String(5, "　")) + ciColGap
      PLeft(6) = PLeft(5) + Printer.TextWidth(String(5, "　")) + ciColGap
   End If
   
End Sub

Private Sub GetPleft2()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   ReDim PLeft(1 To 2)
   PLeft(1) = ciStartX
   PLeft(2) = PLeft(1) + Printer.TextWidth(String(8, "　")) + ciColGap
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      PrintLine
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Private Sub PrintLine()
   Dim iNo As Integer
   iNo = (Printer.ScaleWidth - Printer.CurrentX - 500) \ Printer.TextWidth("-")
   Printer.Print String(iNo, "-")
End Sub


Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "電子機關來函清單"
   Select Case m_strRptSys
   Case "P"
      strPTmp = strPTmp & "-內專"
   Case "T"
      strPTmp = strPTmp & "-內商"
   Case "FCT"
      strPTmp = strPTmp & "-外商"
   Case "FCP"
      strPTmp = strPTmp & "-FCP"
   End Select
   
   'Modified by Morgan 2021/11/30 +And m_bolAllCase = False
   If m_strRptSys = "FCP" And m_strIUser = "" And m_bolAllCase = False Then
      If m_bolNoFileCase = True Then
         strPTmp = strPTmp & "(不調卷)"
      Else
         strPTmp = strPTmp & "(要調卷)"
      End If
   End If
   
   If m_strIUser <> "" Then
      strExc(0) = "select st02 from staff where st01='" & m_strIUser & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         strPTmp = strPTmp & "(" & RsTemp(0) & ")"
      End If
   End If
   
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   
   If m_bolDivList Then
      strPTmp = "分割子案清單"
      iPrint = iPrint + 100
      Printer.Font.Size = 16
      Printer.Font.Bold = True
      Printer.Font.Underline = True
      Printer.CurrentX = lngPageWidth / 2 - Printer.TextWidth(strPTmp) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
      iPrint = iPrint + 200
   End If
   
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False

   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   
   strPTmp = "簽收日期：" & ChangeTStringToTDateString(m_strRptDate)
   Printer.CurrentX = lngPageWidth / 2 - Printer.TextWidth(strPTmp) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   
   PrintNewLine
   
   If m_strRptSys = "P" Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print "＊E化案件"
   End If
   
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
      
   If m_strRptSys = "FCP" Then
      Printer.FontSize = 11
      PrintNewLine
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      'Modified by Morgan 2020/2/24 +◇
      Printer.Print "本所案號後有 ◆ 為不調卷來函  ◇ 為來函需通知承辦寄代"
      PrintNewLine
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print "處理期限前有 * 號者為先正達有期限之來函或Lerner+泰斯拉公司及英帆薩斯公司+Y53942 Xperi 的案件"
      PrintNewLine
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      'Modify By Sindy 2017/10/20
      'Printer.Print "　　　　　有 ** 號者為 DOW 有期限或為 Sandvik OA 需2日內報告之來函"
      Printer.Print "　　　　　有 ** 號者為 DOW 有期限或為 需2日內報告之來函"
      '2017/10/20 END
      PrintNewLine
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print "　　　　　有 *** 號者為UNIUS或YASUTOMI之有期限案件"
      Printer.FontSize = ciFontSize
   End If
   
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   PrintLine
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    For intI = 1 To UBound(PLeft)
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print PColName(intI)
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = 1 To UBound(PLeft)
      Printer.CurrentX = PLeft(iCol)
      Printer.CurrentY = iPrint
      If PColName(iCol) = "案由" Then
         Printer.Print convForm(strData(iCol), 24)
      Else
         Printer.Print strData(iCol)
      End If
    Next
End Sub

'列印表尾
'Modified by Morgan 2020/2/4 +iPaperCount,iNoPaperCount
Private Sub PrintReportFooter(ByVal iRecCount As Integer, Optional iCaseCount As Integer = 0, Optional iPaperCount As Integer = -1, Optional iNoPaperCount As Integer = -1)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    PrintLine
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    If m_bolDivList Then
      Printer.Print "子案數：" & iCaseCount
    Else
      Printer.Print IIf(iCaseCount > 0, "本所案號：" & iCaseCount, "") & vbTab & vbTab & IIf(iRecCount > 0, "公文：" & iRecCount, "") & IIf(iPaperCount >= 0, vbTab & vbTab & "調卷:" & iPaperCount, "") & IIf(iNoPaperCount >= 0, vbTab & vbTab & "◆:" & iNoPaperCount, "")
   End If
End Sub

Private Sub sbAddList(pMessage As String, Optional pAddSpaceLine As Boolean = False, Optional pNoNewItem As Boolean = False)
   Dim stMsg As String
   
   If pAddSpaceLine Then
      lstHistory.AddItem "", 0
      'WriteLog ""
      WriteLog2 ""
   End If
   
   If pNoNewItem Then
      stMsg = "..." & pMessage
      lstHistory.List(0) = lstHistory.List(0) & stMsg
      WriteLog2 stMsg, False
   Else
      stMsg = Now & "   " & pMessage
      lstHistory.AddItem stMsg, 0
      WriteLog2 stMsg
   End If
   
   'WriteLog stMsg
   DoEvents
End Sub

'Added by Morgan 2017/9/1
Private Sub WriteLog2(pLog As String, Optional pNewLine As Boolean = True)
   Dim fs As Object
   Dim stLogFolder As String, stLogFile As String
   
On Error GoTo ErrHnd

   stLogFolder = App.path & "\EDocLog"
   If Dir(stLogFolder, vbDirectory) = "" Then
      MkDir stLogFolder
   End If

   'log保留一年(清除前一年的log)
   'Modifiedby Morgan 2018/11/12 改一個月一次
   'stLogFile = stLogFolder & "\EDoc" & (Format(Now, "yyyyww") - 100) & ".log"
   stLogFile = stLogFolder & "\EDoc" & (Format(Now, "yyyymm") - 100) & ".log"
   If Dir(stLogFile) <> "" Then Kill stLogFile
   
   'stLogFile = stLogFolder & "\EDoc" & Format(Now, "yyyyww") & ".log"
   stLogFile = stLogFolder & "\EDoc" & Format(Now, "yyyymm") & ".log"
   Set fs = CreateObject("Scripting.FileSystemObject")
   fs.OpenTextFile(stLogFile, 8, True).Write IIf(pNewLine, vbCrLf, "") & pLog
   
ErrHnd:
   Set fs = Nothing
End Sub

Private Sub Timer1_Timer()
   lblCountDown = lblCountDown - 1
   If lblCountDown < 1 Then
      Timer1.Enabled = False
      fraCountDown.Visible = False
      m_bolAutoUnload = True 'Added by Morgan 2017/12/5
      cmdAutoRun.Value = True
   End If
End Sub

Private Sub Timer2_Timer()
   If Me.Visible = True Then
      Timer2.Enabled = False
      Form_Activate
   End If
End Sub

Private Sub UpdateOurRef(pED01 As String, pED02 As String, pED10 As String, pED16 As String, pED23 As String, Optional pED15 As String)
   Dim stSQL As String, intR As Integer
   
   '專利
   If pED10 = "P" Then
      '申請號
      If InStr(pED02, "N") = 0 Then
         '申請號
         stSQL = "update edocument set ed27=(select max(pa01||pa02||pa03||pa04) from patent where pa11=ed02 and pa01 in('P','FCP') and pa09='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
         cnnConnection.Execute stSQL, intR
         'Added by Morgan 2018/1/16 重新申請若未收新案號時舊申請號手動放CP30以便輸入來函 Ex.106144935(FCP-058047)
         '對方案號
         stSQL = "update edocument set ed27=(select max(pa01||pa02||pa03||pa04) from caseprogress,patent where cp30=ed02 and cp01 in ('P','FCP') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
         cnnConnection.Execute stSQL, intR
         'end 2018/1/16
      '舉發案號
      Else
         '對造號
         stSQL = "update edocument set ed27=(select max(pa01||pa02||pa03||pa04) from caseprogress,patent where cp36=ed02 and cp01 in ('P','FCP') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
         cnnConnection.Execute stSQL, intR
         '申請號
         stSQL = "update edocument set ed27=(select max(pa01||pa02||pa03||pa04) from patent where pa11=substr(ed02,1,9) and pa01 in('P','FCP') and pa09='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
         cnnConnection.Execute stSQL, intR
      End If
   '商標
   Else
      'Added by Morgan 2022/6/29 商標一文多案申請號及註冊號的順序可能不一致,故改判斷非爭議號時統一用申請案號檢查
      If pED15 = "" Then
         stSQL = "update edocument set ed15=ed15 where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null and ed15='爭議號'"
         cnnConnection.Execute stSQL, intR
         If intR = 1 Then
            pED15 = "爭議號"
         End If
      End If
      If pED15 <> "爭議號" Then
         '申請號
         If pED02 <> "" Then
            'Added by Morgan 2023/7/25 未閉卷有A類申請101的優先 Ex:T-244689,FCT-050092
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from trademark where tm12=ed02 and tm01 in('T','FCT') and tm10='000'" & _
            " and tm57 is null and exists(select * from caseprogress where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp10='101' and cp09<'B')" & _
            ") where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
            'end 2023/7/25
            
            'Added by Morgan 2022/7/11 Ex:FCT038473000
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from trademark where tm12=ed02 and tm01 in('T','FCT') and tm10='000' and (tm15 is null or (tm57 is null and tm16='1'))) where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
            'end 2022/7/11
            
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from trademark where tm12=ed02 and tm01 in('T','FCT') and tm10='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
            
         End If
      Else
      'end 2022/6/29
      
         '註冊號
         If pED16 <> "" Then
            'Modified by Morgan 2017/6/22 排除已銷卷案
            '加已核准條件
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from trademark where tm15=ed16 and tm01 in('T','FCT') and tm10='000' and tm57 is null and tm16='1') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
         End If
         '申請號
         If pED02 <> "" Then
         
            'Added by Morgan 2023/7/25 未閉卷有A類申請101的優先 Ex:T-244689,FCT-050092
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from trademark where tm12=ed02 and tm01 in('T','FCT') and tm10='000'" & _
            " and tm57 is null and exists(select * from caseprogress where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and cp10='101' and cp09<'B')" & _
            ") where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
            'end 2023/7/25
            
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from trademark where tm12=ed02 and tm01 in('T','FCT') and tm10='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
            '對方案號
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from caseprogress,trademark where cp30=ed02 and cp01 in('T','FCT') and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm10='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
         End If
         'Modified by Morgan 2017/11/16 自上面移下來改先抓申請號,Ex.103062533 =>FCT040394的對造號(清單),FCT040393的申請號(畫面)
         'Added by Morgan 2017/7/4 對造會放審定號 FCT-039231 (基本資料多存了空白但對造是正確的)
         '對造號
         If pED16 <> "" Then
            stSQL = "update edocument set ed27=(select max(tm01||tm02||tm03||tm04) from caseprogress,trademark where cp36=ed16 and cp01 in('T','FCT') and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm10='000') where ed01='" & pED01 & "' and ed23='" & pED23 & "' and ed27 is null"
            cnnConnection.Execute stSQL, intR
         End If
         
      End If
   End If
End Sub

Private Function fnChkFile(pFile As String) As Boolean
On Error GoTo ErrHnd
   If Dir(pFile) <> "" Then
      fnChkFile = True
   End If
ErrHnd:
End Function

'Added by Morgan 2020/2/21
'Modified by Morgan 2020/6/17 +pCP10
'不調卷例外設定
Private Function fnFCPXCase(pCP10 As String, pYNo As String, pXNo As String) As Boolean
   If pCP10 = "1204" Then 'Added by Morgan 2020/6/17
      If InStr("X45149000,X60049C10", pXNo) > 0 Then
         fnFCPXCase = True
      'Modified by Morgan 2020/4/10 +Y54339,Y54339B1,Y54339B2 --Jessica,Lisa
      ElseIf pYNo = "Y20624000" Or pYNo = "Y54339000" Or pYNo = "Y54339B10" Or pYNo = "Y54339B20" Then
         fnFCPXCase = True
      ElseIf pYNo = "Y54116000" And pXNo = "X48637000" Then
         fnFCPXCase = True
      ElseIf pYNo = "Y34232000" And pXNo = "X48637000" Then
         fnFCPXCase = True
      End If
   End If
End Function

'Added by Morgan 2023/5/29
'更新商標一文多案為主案的申請號(ED02)/註冊號(ED16)/本所案號(ED27)
Private Sub UpdMultiDoc(pED01 As String, pED23 As String)
   Dim strQ As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   strQ = "select d.*  From edocument,trademark a,caseprogress b,caseprogress c,trademark d" & _
      " where ed01='" & pED01 & "' and ed23='" & pED23 & "'" & _
      " and a.tm01(+) = substr(ed27,1,length(ed27)-9)" & _
      " and a.tm02(+) = substr(ed27,-9,6)" & _
      " and a.tm03(+) = substr(ed27,-3,1)" & _
      " and a.tm04(+) = substr(ed27,-2)" & _
      " and b.cp01(+)=a.tm01 and b.cp02(+)=a.tm02 and b.cp03(+)=a.tm03 and b.cp04(+)=a.tm04 and b.cp10 in('301','501')" & _
      " and c.cp09(+)=b.cp28 and d.tm01(+)=c.cp01 and d.tm02(+)=c.cp02 and d.tm03(+)=c.cp03 and d.tm04(+)=c.cp04 and d.tm12 is not null" & _
      " order by c.cp27 desc"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      With rsQuery
      strQ = "update edocument set (ed02,ed16,ed27)=(select tm12,nvl(tm15,tm12),tm01||tm02||tm03||tm04" & _
         " from trademark where tm01='" & .Fields("tm01") & "' and tm02='" & .Fields("tm02") & "'" & _
         " and tm03='" & .Fields("tm03") & "' and tm04='" & .Fields("tm04") & "')" & _
         " where ed01='" & pED01 & "' and ed23='" & pED23 & "'"
      cnnConnection.Execute strQ, intQ
      End With
   End If
   
   Set rsQuery = Nothing
End Sub

'Added by Morgan 2024/1/18
'若副檔名不是pdf時補.pdf(新版的reader會不列印非.pdf的檔案)
Private Sub SetPdfFile(ByRef pFile As String)
   If LCase(Right(pFile, 4)) <> ".pdf" Then
      Name pFile As pFile & ".pdf"
      pFile = pFile & ".pdf"
   End If
End Sub
