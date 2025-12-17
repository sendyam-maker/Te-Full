VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmAutoPdf 
   AutoRedraw      =   -1  'True
   Caption         =   "定稿轉PDF"
   ClientHeight    =   4692
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   7596
   Icon            =   "frmAutoPdf.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4692
   ScaleWidth      =   7596
   Begin VB.FileListBox File2 
      Height          =   252
      Left            =   1080
      TabIndex        =   10
      Top             =   3840
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrTFA 
      Left            =   6180
      Top             =   3960
   End
   Begin VB.Timer tmrMoveList 
      Left            =   3900
      Top             =   3960
   End
   Begin VB.TextBox TxtFile 
      Height          =   285
      Left            =   4380
      TabIndex        =   9
      Text            =   "TxtFile"
      Top             =   4020
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CheckBox ChkBox1 
      Caption         =   "要執行電子發票上傳"
      Height          =   225
      Left            =   60
      TabIndex        =   8
      Top             =   4050
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   252
      Left            =   96
      TabIndex        =   7
      Top             =   3816
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Timer tmrTranInv 
      Left            =   2100
      Top             =   90
   End
   Begin VB.Timer tmrMail 
      Left            =   810
      Top             =   90
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   2130
      _ExtentX        =   593
      _ExtentY        =   593
      _Version        =   393216
   End
   Begin VB.CommandButton cmdConnect 
      Caption         =   "連線"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   725
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "停止"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4950
      TabIndex        =   5
      Top             =   90
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Height          =   345
      Left            =   2520
      ScaleHeight     =   300
      ScaleWidth      =   900
      TabIndex        =   4
      Top             =   120
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.ListBox lstHistory 
      Height          =   3108
      Left            =   45
      TabIndex        =   2
      Top             =   630
      Width           =   7455
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "開始"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3600
      TabIndex        =   1
      Top             =   90
      Width           =   1275
   End
   Begin VB.Timer tmrPolling 
      Left            =   1245
      Top             =   90
   End
   Begin VB.Timer tmrClock 
      Left            =   1665
      Top             =   90
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6300
      TabIndex        =   0
      Top             =   90
      Width           =   1050
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  '對齊表單下方
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Top             =   4380
      Width           =   7590
      _ExtentX        =   13399
      _ExtentY        =   550
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2970
      Top             =   3930
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuShow 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu mnuDisplay 
         Caption         =   "顯示"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "結束"
      End
   End
End
Attribute VB_Name = "frmAutoPdf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/2/22 改成Form2.0 (無)
'Created by Morgan 2014/3/27
Option Explicit

Dim bolActived As Boolean
Dim mlngID As Long
Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim stToPath As String 'Add by Amy 2019/04/19
'Add by Amy 2019/07/11 零晨2點左右會斷線,有些程式不需一直執行
Const strTimeS = "0700" '起始時間
Const strTimeE = "1900" '結束時間
Dim strRunExePC As String 'Add by Amy 2019/12/27
Dim bolMoveStatus As Boolean 'Added by Lydia 2020/03/19 是否正在執行搬檔作業中

'Add by Amy 2019/12/27
Private Sub ChkBox1_Click()
    Dim stMsg As String
    '測式資料庫,連線至測式平台,彈訊息告知開啟 Amy電腦
    If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 And ChkBox1.Value = 1 Then
        stMsg = "需連線至盟立測式平台,請確認：" & vbCrLf & _
                      "1.產生之XML檔需為測式Tag" & vbCrLf & _
                      "2.盟立平台是否已新增字軌" & vbCrLf & _
                      "3.程式需在Amy電腦執行,請確認Amy電腦是否開啟" & vbCrLf & vbCrLf & _
                      "以上確認無誤, 請按「是」繼續操作"
        If MsgBox(stMsg, vbYesNo + vbCritical) = vbNo Then
            ChkBox1.Value = 0
            Exit Sub
        'Add by Amy 2024/09/30 只要測式發票上傳
        Else
             cmdStart_Click
        End If
    End If
End Sub
'end 2019/12/27

Private Sub cmdConnect_Click()
   If PUB_Connect2DB(True) = False Then
      lstHistory.AddItem Now & "--> 連線已切換"
      Unload Me
   'Add by Amy 2019/07/11 切換至正式機勾選
   Else
        'ChkBox1.Value = 0 'Mark by Amy 2019/12/27 改至FormLoad
        If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then
           ChkBox1.Value = 1
        End If
   End If
End Sub

Private Sub cmdStart_Click()
   'Add by Amy 2024/09/30 測式用
   If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 And ChkBox1.Value = 1 Then
      StatusBar1.Panels.Item(1).Text = "等待中..."
      cmdStart.Enabled = False
      cmdStop.Enabled = True
      lstHistory.AddItem Now & "--> 開始"
      tmrTranInv.Interval = 2000
      tmrTranInv.Enabled = True
      Exit Sub
   End If
   
   tmrPolling.Interval = 1000
   tmrPolling.Enabled = True
   StatusBar1.Panels.Item(1).Text = "等待中..."
   cmdStart.Enabled = False
   cmdStop.Enabled = True
   lstHistory.AddItem Now & "--> 開始"
   'Added by Lydia 2017/12/12 FCP案件命名通知信
   'Modified by Lydia 2018/03/08 改成1分鐘x5
   tmrMail.Interval = 60000
   tmrMail.Enabled = True
   'end 2017/12/12
   'Added by Lydia 2020/03/19
   tmrMoveList.Interval = 60000
   tmrMoveList.Enabled = True
   'end 2020/03/19
   'Added by Lydia 2023/05/05
   tmrTFA.Interval = 60000
   tmrTFA.Enabled = True
   'end 2023/05/05
   
   'Add by Amy 2019/04/19 勾選「要執行電子發票上傳」才上傳
   '選正式資料庫才被勾選,會上傳至廠商正式機(設定上傳於廠商正式機在C:\551cron\set.ini)
   If Val(Format(Now, "YYYYMMDD")) >= Val(TranInvoiceDate) + 19110000 Then
        If ChkBox1.Value = 1 Then
             tmrTranInv.Interval = 60000 'Modify by Amy 2019/07/11 改1分鐘 原:2000
             tmrTranInv.Enabled = True
        End If
   End If
End Sub

Private Sub cmdExit_Click()
   lstHistory.AddItem Now & "--> 程式結束"
   Unload Me
End Sub

Private Sub cmdStop_Click()
   tmrPolling.Enabled = False
   tmrPolling.Interval = 0
   StatusBar1.Panels.Item(1).Text = ""
   cmdStart.Enabled = True
   cmdStop.Enabled = False
   lstHistory.AddItem Now & "--> 轉檔停止"
   'Added by Lydia 2017/12/12 FCP案件命名通知信
   tmrMail.Enabled = False
   tmrMail.Interval = 0
   'end 2017/12/12
   'Added by Lydia 2020/03/19
   tmrMoveList.Interval = 0
   tmrMoveList.Enabled = False
   'end 2020/03/19
   'Add by Amy 2019/04/19 勾選「要執行電子發票上傳」才上傳
   '選正式資料庫才被勾選,會上傳至廠商正式機(設定上傳於廠商正式機在C:\551cron\set.ini)
   If ChkBox1.Value = 1 Then
        tmrTranInv.Enabled = False
        tmrTranInv.Interval = 0
   End If
End Sub

Private Sub Form_Activate()
   Screen.MousePointer = vbHourglass
   If bolActived = False Then
      Me.Tag = Me.Caption
      Me.Top = (Screen.Height - Me.Height) / 2
      Me.Left = (Screen.Width - Me.Width) / 2
      If fConnect(Me) = False Then
         Unload Me
      Else
         bolMailFailNoAlert = True
         '關閉鈕 鎖 x 變灰色
         DisableControl Me
      End If
      
      'Add by Amy 2019/07/11
      'ChkBox1.Value = 0 'Mark by Amy 2019/12/27
      If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then
         ChkBox1.Value = 1
      End If
      'end 2019/07/11
      
      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
         cmdStart.Value = True
      End If

      bolActived = True
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   'Add by Amy 2024/10/23
   If CheckIsRunning(App.EXEName & ".exe") Then
      MsgBox "程式已正在執行中，不可重複執行...", vbExclamation
      End
   End If
   
   stToPath = "C:\Einvoice\" 'Add by Amy 2019/04/19 發票上傳用
   ChkBox1.Value = 0 'Add by Amy 2019/12/27 由cmdConnect_Click搬來,連測式機資料庫時可以連盟立測式平台
   tmrClock.Interval = 1000
   lstHistory.Clear
   lstHistory.AddItem Now & "--> 程式已載入"
   If mlngID = 0 Then mlngID = AddToSystemTray(Picture1.hWnd, WM_MOUSEMOVE, Me.Icon, Me.Caption)
End Sub

Private Sub Form_Resize()
If Me.WindowState = "1" Then Me.Visible = False
End Sub

'Add By Sindy 2019/11/18
Public Sub KillFile()
   If Dir(App.path & "\$$*.pdf") <> "" Then
      Kill App.path & "\$$*.pdf"
   End If
   If Dir(App.path & "\$$*.jpg") <> "" Then
      Kill App.path & "\$$*.jpg"
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set oFileSys = Nothing
   Set oFile = Nothing
   
   'Add By Sindy 2019/11/18
'   If Dir(App.path & "\$$*.pdf") <> "" Then
'      Kill App.path & "\$$*.pdf"
'   End If
   KillFile
   '2019/11/18 END
   
   If mlngID <> 0 Then
      DeleteFromSystemTray mlngID
      mlngID = 0
   End If
   
   WriteLog True
   Set frmAutoPdf = Nothing
End Sub

Private Sub mnuDisplay_Click()
Me.WindowState = "0"
Me.Visible = True
End Sub

Private Sub mnuQuit_Click()
Unload Me
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim MSG As Long

If Me.ScaleMode = 1 Then
   MSG = x / Screen.TwipsPerPixelX
End If
Select Case MSG
      Case WM_MOUSEMOVE '移動滑鼠
          'Label1.Caption = "正在移動滑鼠"
      Case WM_LBUTTONDBLCLK '連點滑鼠左鍵
          'Label1.Caption = "連點滑鼠左鍵"
          Me.WindowState = "0"
          Me.Visible = True
      Case WM_LBUTTONDOWN '按下滑鼠左鍵
          'Label1.Caption = "按下滑鼠左鍵"
      Case WM_LBUTTONUP '放開滑鼠左鍵
          'Label1.Caption = "放開滑鼠左鍵"
      Case WM_RBUTTONDBLCLK '連點滑鼠右鍵
          'Label1.Caption = "連點滑鼠右鍵"
      Case WM_RBUTTONDOWN '按下滑鼠右鍵
          'Label1.Caption = "按下滑鼠右鍵"
          Me.PopupMenu mnuShow, vbPopupMenuLeftAlign + vbPopupMenuRightButton
      Case WM_RBUTTONUP '放開滑鼠右鍵
          ''Label1.Caption = "放開滑鼠右鍵"
End Select
End Sub
'Modified by Morgan 2016/12/30 改 NT2(live),M51-1(test)都要檢查
Private Sub tmrClock_Timer()
   Static dtlastTimeLive As Date
   Static dtlastTimeTest As Date
   Static iMinLive As Integer
   Static iMinTest As Integer
   Static stDate As String
   Static bolInformLive As Boolean
   Static bolInformTest As Boolean
   Static iCountLive As Integer
   Static iCountTest As Integer
   Dim bolConnErr As Boolean
   Dim stDbName As String
      
   StatusBar1.Panels.Item(2).Text = Time
   
   Exit Sub 'Added by Morgan 2018/6/20 改用 O12 先取消
   
   If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then Exit Sub 'Added by Morgan 2017/9/15
   
   If stDate <> Format(Now, "YYYYMMDD") Then
      stDate = Format(Now, "YYYYMMDD")
      bolInformLive = False
      bolInformTest = False
   End If
   
   'NT2 例行通知:6點(周六、日除外)
   bolConnErr = False
   If Weekday(Now) > 1 And Weekday(Now) < 7 Then
      If Not bolInformLive Then
         If Val(Format(Now, "HH")) >= 6 And Val(Format(Now, "HH")) <= 7 Then
            bolInformLive = True
            If TableSpaceInformNew("live", , , stDbName) = False Then
               PUB_SendMail "QPGMR", "74001;83002;92012", "", "(" & stDbName & ") Table Space 統計失敗！", "如旨", , , , , , , "QPGMR", , , , False
            End If
         End If
      End If
   End If
   
   'NT2 異常通知:每3分鐘檢查1次,發現異常後每30分鐘檢查1次(已發信原則應該有人處理,一直發信也沒用,30分鐘發1次觀察變化量)
   If Not bolConnErr Then
      If Now > dtlastTimeLive + iMinLive / (24 * 60) Then
         dtlastTimeLive = Now
         
         If TableSpaceInformNew("live", True, bolConnErr, stDbName) Then
            iMinLive = 30
            
         ElseIf bolConnErr Then
            PUB_SendMail "QPGMR", "74001;83002;92012", "", "(" & stDbName & ") 連線失敗！", "如旨", , , , , , , "QPGMR", , , , False
            iMinLive = 60
            
         Else
            iMinLive = 3
         End If
      End If
   End If
      
'Removed by Morgan 2017/8/31 取消,測試改用O12但目前無法與O8切換連線
'   'M51 (週一~週五8點以後)
'   If (Weekday(Now) > 1 And Weekday(Now) < 7) And Val(Format(Now, "HH")) >= 8 Then
'      'M51 例行通知:22點
'      bolConnErr = False
'      If Not bolInformTest Then
'         If Val(Format(Now, "HH")) >= 22 And Val(Format(Now, "HH")) <= 23 Then
'            bolInformTest = True
'
'            If TableSpaceInformNew("test", , , stDbName) = False Then
'               PUB_SendMail "QPGMR", "74001;83002;92012", "", "(" & stDbName & ") Table Space 統計失敗！", "如旨", , , , , , , "QPGMR", , , , False
'            End If
'         End If
'      End If
'
'      'M51 每3分鐘檢查1次,發現異常後每30分鐘檢查1次
'      If Not bolConnErr Then
'         If Now > dtlastTimeTest + iMinTest / (24 * 60) Then
'            dtlastTimeTest = Now
'            If TableSpaceInformNew("test", True, bolConnErr) Then
'               iMinTest = 30
'
'            ElseIf bolConnErr Then
'               If Val(Format(Now, "hhmm")) < 800 Then
'                  iMinTest = 830 - Val(Format(Now, "hhmm"))
'               Else
'                  iMinTest = 60
'               End If
'            Else
'               iMinTest = 3
'            End If
'         End If
'      End If
'   End If
'end 2017/8/31

End Sub

'Memo by Morgan 2016/12/30 改用TableSpaceInformNew
Private Function TableSpaceAlarm() As Boolean
   Dim stSQL As String, intQ As Integer
   Dim stSubject As String, stContent As String
   Dim rsQuery As ADODB.Recordset
   Dim stMinSize As String

   stMinSize = "3*1024*1024*1024"

   stSQL = "SELECT df.TABLESPACE_NAME, MIN(df.BYTES-nvl(fs.BYTES,0)) Used_Bytes" & _
      " from dba_data_files df,(select TABLESPACE_NAME,FILE_ID,sum(fs.BYTES) bytes" & _
      " from dba_free_space fs group by TABLESPACE_NAME,FILE_ID ) fs" & _
      " where fs.TABLESPACE_NAME(+)=df.TABLESPACE_NAME" & _
      " and fs.FILE_ID(+)=df.FILE_ID" & _
      " GROUP BY df.TABLESPACE_NAME" & _
      " HAVING MIN(df.BYTES-nvl(fs.BYTES,0))>" & stMinSize
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If rsQuery.RecordCount = 1 Then
         stSubject = "注意:Table Space【 " & rsQuery(0) & " 】已達警戒值!!!(最小的已使用空間超過3GB)"
         stContent = "Table Space:" & rsQuery(0) & vbTab & " 最小的已使用空間:" & Format(rsQuery(1), "#,##0") & " Bytes"
      Else
         stSubject = "注意:有多個 Table Space 已達警戒值!!!(最小的已使用空間超過3GB)"
         stContent = "Table Space:" & rsQuery(0) & vbTab & " 最小的已使用空間:" & Format(rsQuery(1), "#,##0") & " Bytes"
         rsQuery.MoveNext
         Do While Not rsQuery.EOF
            stContent = stContent & vbCrLf & "Table Space:" & rsQuery(0) & vbTab & " 最小的已使用空間:" & Format(rsQuery(1), "#,##0") & " Bytes"
            rsQuery.MoveNext
         Loop
      End If

      'Modify By Sindy 2016/7/12 bolCaseDutyAgentMsg=False,才不會收件者休假彈出寄信給職代的訊息
      'PUB_SendMail "QPGMR", "74001;83002;92012", "", stSubject, stContent, , , , , , , "QPGMR"
      PUB_SendMail "QPGMR", "74001;83002;92012", "", stSubject, stContent, , , , , , , "QPGMR", , , , False
      TableSpaceAlarm = True
   End If

   Set rsQuery = Nothing
End Function

'Memo by Morgan 2016/12/30 改用TableSpaceInformNew
Private Function TableSpaceInform() As Boolean
   Dim stSQL As String, intQ As Integer
   Dim stSubject As String, stContent As String
   Dim rsQuery As ADODB.Recordset
   Dim strSpaceName As String
   Dim strVTB As String

   '每一 table space 最小的已使用空間的 file id
   strVTB = "SELECT df.TABLESPACE_NAME, to_number(substr(MIN(100*(df.BYTES-nvl(fs.BYTES,0))+df.FILE_ID),-2)) FILE_ID" & _
      " from dba_data_files df,(select TABLESPACE_NAME,FILE_ID,sum(fs.BYTES) bytes" & _
      " from dba_free_space fs group by TABLESPACE_NAME,FILE_ID ) fs" & _
      " where fs.TABLESPACE_NAME(+)=df.TABLESPACE_NAME" & _
      " and fs.FILE_ID(+)=df.FILE_ID" & _
      " GROUP BY df.TABLESPACE_NAME"

   stSQL = "SELECT df.FILE_NAME, df.TABLESPACE_NAME, round(df.BYTES/(1024*1024),3) Bytes, round((df.BYTES-nvl(fs.BYTES,0))/(1024*1024),3) Used_Bytes,round(nvl(fs.BYTES,0)/(1024*1024),3) Free_Bytes,df.FILE_ID,nvl(m.FILE_ID,0) MIN_FILE_ID" & _
      " from dba_data_files df,(select TABLESPACE_NAME,FILE_ID,sum(fs.BYTES) bytes" & _
      " from dba_free_space fs group by TABLESPACE_NAME,FILE_ID ) fs,(" & strVTB & ") m" & _
      " where fs.TABLESPACE_NAME(+)=df.TABLESPACE_NAME and fs.FILE_ID(+)=df.FILE_ID" & _
      " and m.TABLESPACE_NAME(+)=df.TABLESPACE_NAME and m.FILE_ID(+)=df.FILE_ID" & _
      " order by 2,1"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stSubject = "Table Space 統計"
      strSpaceName = rsQuery("TABLESPACE_NAME")

      rsQuery.MoveFirst
      '表頭
      stContent = "<BODY>" & vbCrLf & "<BR>&nbsp;" & vbCrLf & "<TABLE BORDER CELLSPACING=1 CELLPADDING=1 WIDTH=550>"
      stContent = stContent & vbCrLf & "<TR><TD WIDTH=""36%"" VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Name</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD WIDTH=""24%"" VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Tablespace</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD WIDTH=""20%"" VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Size(M)</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD WIDTH=""20%"" VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Used(M)</FONT></TD>"

      '第1筆
      stContent = stContent & vbCrLf & "<TR><TD VALIGN=""TOP"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & rsQuery("FILE_NAME") & "</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & rsQuery("TABLESPACE_NAME") & "</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(rsQuery("Bytes"), "#,##0.000") & "</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20 " & IIf(rsQuery("FILE_ID") = rsQuery("MIN_FILE_ID"), IIf(rsQuery("Used_Bytes") > 2000, "BGCOLOR=""#ff0000""", "BGCOLOR=""#ffff00"""), "") & ">"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(rsQuery("Used_Bytes"), "#,##0.000") & "</FONT></TD></TR>"

      rsQuery.MoveNext
      Do While Not rsQuery.EOF
         stContent = stContent & vbCrLf & "<TR><TD VALIGN=""TOP"" HEIGHT=20>"
         stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & rsQuery("FILE_NAME") & "</FONT></TD>"
         stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
         stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & rsQuery("TABLESPACE_NAME") & "</FONT></TD>"
         stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
         stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(rsQuery("Bytes"), "#,##0.000") & "</FONT></TD>"
         stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20 " & IIf(rsQuery("FILE_ID") = rsQuery("MIN_FILE_ID"), IIf(rsQuery("Used_Bytes") > 2000, "BGCOLOR=""#ff0000""", "BGCOLOR=""#ffff00"""), "") & ">"
         stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(rsQuery("Used_Bytes"), "#,##0.000") & "</FONT></TD></TR>"
         rsQuery.MoveNext
      Loop
      stContent = stContent & vbCrLf & "</TABLE>" & vbCrLf & "</BODY>"

      'Modify By Sindy 2016/7/12 bolCaseDutyAgentMsg=False,才不會收件者休假彈出寄信給職代的訊息
      'PUB_SendMail "QPGMR", "74001;83002;92012", "", stSubject, stContent, , , , , , , "QPGMR"
      PUB_SendMail "QPGMR", "74001;83002;92012", "", stSubject, stContent, , , , , , , "QPGMR", , , , False
      TableSpaceInform = True
   End If

   Set rsQuery = Nothing
End Function

'Added by Morgan 2016/12/30
Private Function TableSpaceInformNew(pDS As String, Optional pAlarm As Boolean = False, Optional pConnErr As Boolean, Optional pDbName As String) As Boolean
   
   Dim cnnQurey As New ADODB.Connection
   Dim rsQuery As New ADODB.Recordset
   Dim strVTB As String, stSQL As String, intQ As Integer
   Dim stDB As String, stSubject As String, stContent As String
   Dim stMinSize As String, stSysMinSize As String
   Dim iRecs As Integer
   Dim bolAdd As Boolean, stAlert As String
   Dim bolConnErr As Boolean
   Dim strSQLContext As String
   Dim strLog As String
   
   bolConnErr = True
   pDbName = ""
   
On Error GoTo ErrConn

   WriteLog2 "Start: " & IIf(pAlarm, "Alarm", "Regular"), pDS

   If cnnQurey.State = adStateClosed Then
      cnnQurey.ConnectionString = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=" & pDS & ";Persist Security Info=True"
      cnnQurey.Open
   End If
   
   bolConnErr = False
   
   stSQL = "select TERMINAL FROM V$SESSION where SID=1"
   If rsQuery.State = adStateOpen Then rsQuery.Close
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnQurey, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      stDB = "" & rsQuery(0)
      pDbName = stDB
   End If

   stMinSize = "3" 'GB
   stSysMinSize = "1" 'GB
   
   '每一 table space 最小的已使用空間的 file id
   strVTB = "SELECT df.TABLESPACE_NAME, to_number(substr(MIN(100*(df.BYTES-nvl(fs.BYTES,0))+df.FILE_ID),-2)) FILE_ID" & _
      " from dba_data_files df,(select TABLESPACE_NAME,FILE_ID,sum(fs.BYTES) bytes" & _
      " from dba_free_space fs group by TABLESPACE_NAME,FILE_ID ) fs" & _
      " where fs.TABLESPACE_NAME(+)=df.TABLESPACE_NAME" & _
      " and fs.FILE_ID(+)=df.FILE_ID" & _
      " GROUP BY df.TABLESPACE_NAME"
   
   stSQL = "SELECT df.FILE_NAME, df.TABLESPACE_NAME, round(df.BYTES/(1024*1024),3) Bytes, round((df.BYTES-nvl(fs.BYTES,0))/(1024*1024),3) Used_Bytes,round(nvl(fs.BYTES,0)/(1024*1024),3) Free_Bytes,df.FILE_ID,nvl(m.FILE_ID,0) MIN_FILE_ID" & _
      " from dba_data_files df,(select TABLESPACE_NAME,FILE_ID,sum(fs.BYTES) bytes" & _
      " from dba_free_space fs group by TABLESPACE_NAME,FILE_ID ) fs,(" & strVTB & ") m" & _
      " where fs.TABLESPACE_NAME(+)=df.TABLESPACE_NAME and fs.FILE_ID(+)=df.FILE_ID" & _
      " and m.TABLESPACE_NAME(+)=df.TABLESPACE_NAME and m.FILE_ID(+)=df.FILE_ID" & _
      " order by 2,1"
   
   If rsQuery.State = adStateOpen Then rsQuery.Close
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnQurey, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      With rsQuery
      .MoveFirst
      '表頭
      stContent = "<BODY>" & vbCrLf & "<BR>&nbsp;" & vbCrLf & "<TABLE BORDER CELLSPACING=1 CELLPADDING=1 WIDTH=550>"
      stContent = stContent & vbCrLf & "<TR><TD WIDTH=""35%"" VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Name</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD WIDTH=""20%"" VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Tablespace</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD WIDTH=""16%"" VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Size(M)</FONT></TD>"
      stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" BGCOLOR=""#c0c0c0"" HEIGHT=20>"
      stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>Used(M)</FONT></TD>"
      
      If pAlarm Then
         Do While Not .EOF
            strLog = strLog & " " & .Fields("TABLESPACE_NAME") & "=" & .Fields("Used_Bytes")
            
            If .Fields("FILE_ID") = .Fields("MIN_FILE_ID") Then
               bolAdd = False
               Select Case UCase(.Fields("TABLESPACE_NAME"))
               Case "RBS", "SYSTEM", "TEMPORARY"
                  If .Fields("Used_Bytes") > Val(stSysMinSize) * 1024 Then
                     iRecs = iRecs + 1
                     bolAdd = True
                     stAlert = " (超過" & stSysMinSize & "GB)"
                     If iRecs = 1 Then
                        stSubject = "(" & stDB & ") 注意:Table Space【 " & .Fields("TABLESPACE_NAME") & " 】已達警戒值!!!(最小的已使用空間超過" & stSysMinSize & "GB)"
                     End If
                  End If
               Case Else
                  If .Fields("Used_Bytes") > Val(stMinSize) * 1024 Then
                     iRecs = iRecs + 1
                     bolAdd = True
                     stAlert = " (超過" & stMinSize & "GB)"
                     If iRecs = 1 Then
                        stSubject = "注意:Table Space【 " & .Fields("TABLESPACE_NAME") & " 】已達警戒值!!!(最小的已使用空間超過" & stMinSize & "GB)"
                     End If
                  End If
               End Select
               
               If bolAdd Then
                  stContent = stContent & vbCrLf & "<TR><TD VALIGN=""TOP"" HEIGHT=20>"
                  stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & .Fields("FILE_NAME") & "</FONT></TD>"
                  stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
                  stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & .Fields("TABLESPACE_NAME") & "</FONT></TD>"
                  stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
                  stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(.Fields("Bytes"), "#,##0.000") & "</FONT></TD>"
                  stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20 BGCOLOR=""#ff0000"">"
                  stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(.Fields("Used_Bytes"), "#,##0.000") & stAlert & "</FONT></TD></TR>"
               End If
            End If
            .MoveNext
         Loop
         
         If iRecs > 1 Then
            stSubject = "(" & stDB & ") 注意:有多個 Table Space 已達警戒值!!!"
         End If
         
      Else
         stSubject = "(" & stDB & ") Table Space 統計"
         Do While Not .EOF
            strLog = strLog & " " & .Fields("TABLESPACE_NAME") & "=" & .Fields("Used_Bytes")
            
            iRecs = iRecs + 1
            stContent = stContent & vbCrLf & "<TR><TD VALIGN=""TOP"" HEIGHT=20>"
            stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & .Fields("FILE_NAME") & "</FONT></TD>"
            stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
            stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P>" & .Fields("TABLESPACE_NAME") & "</FONT></TD>"
            stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20>"
            stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(.Fields("Bytes"), "#,##0.000") & "</FONT></TD>"

            stContent = stContent & vbCrLf & "<TD VALIGN=""TOP"" HEIGHT=20 " & IIf(.Fields("FILE_ID") = .Fields("MIN_FILE_ID"), IIf(.Fields("Used_Bytes") > stMinSize * 1000, "BGCOLOR=""#ff0000""", "BGCOLOR=""#ffff00"""), "") & ">"
            stContent = stContent & vbCrLf & "<FONT FACE=""新細明體"" LANG=""ZH-TW""><P ALIGN=""RIGHT"">" & Format(.Fields("Used_Bytes"), "#,##0.000") & "</FONT></TD></TR>"
            .MoveNext
         Loop
      End If
      
      If pAlarm And iRecs > 0 Then
         stSQL = "SELECT SUBSTRB(A.SID,1,3)||',' SID, SUBSTRB(serial#,1,7) serial#, A.TERMINAL" & _
            ", A.OSUSER, A.PROGRAM APP" & _
            ", B.STATE, DECODE(ST06,'1','北','2','中','3','南','4','高','他') AREA, ST02, A.STATUS" & _
            ",F.PIECE,F.SQL_TEXT" & _
            " FROM V$SESSION A, V$SESSION_WAIT B, V$SQLAREA C, STAFF E, V$SQLTEXT F" & _
            " WHERE A.TYPE='USER' AND NOT EXISTS(SELECT * FROM v$mystat X where rownum=1 AND X.SID=A.SID) AND B.SID=A.SID AND C.ADDRESS(+)= A.SQL_ADDRESS" & _
            " AND STATUS='ACTIVE' AND F.HASH_VALUE=C.HASH_VALUE" & _
            " AND E.ST01(+)=A.OSUSER" & _
            " ORDER BY A.STATUS,ST06,A.SID,F.PIECE"
            
         If .State = adStateOpen Then .Close
         .CursorLocation = adUseClient
         .Open stSQL, cnnQurey, adOpenStatic, adLockReadOnly
         If .RecordCount > 0 Then
            strSQLContext = "<P>---以下為目前執行中的語法---" & "<BR>"
            strSQLContext = strSQLContext & "TERMINAL: " & .Fields("TERMINAL") & "<BR>"
            strSQLContext = strSQLContext & "OSUSER: " & .Fields("OSUSER") & "<BR>"
            strSQLContext = strSQLContext & "APP: " & .Fields("APP") & "<BR>"
            strSQLContext = strSQLContext & "STATE: " & .Fields("STATE") & "<BR>"
            strSQLContext = strSQLContext & "AREA: " & .Fields("AREA") & "<BR>"
            strSQLContext = strSQLContext & "USER: " & .Fields("ST02") & "<BR>"
            strSQLContext = strSQLContext & "SQL: " & "<BR>"
            Do While Not .EOF
               strSQLContext = strSQLContext & .Fields("SQL_TEXT")
               .MoveNext
            Loop
         End If
      End If
         
      If strSQLContext <> "" Then stContent = stContent & "<TR><TD VALIGN=""TOP"" HEIGHT=20 colspan=4>" & strSQLContext & "</TD></TR>"
      
      stContent = stContent & vbCrLf & "</TABLE>" & vbCrLf & "</BODY>"
      If iRecs > 0 Then
         PUB_SendMail "QPGMR", "74001;83002;92012", "", stSubject, stContent, , , , , , , "QPGMR", , , , False
         TableSpaceInformNew = True
      End If
      
      End With
   End If
   
   WriteLog2 strLog, pDS
   
ErrConn:
   If bolConnErr Then
      pConnErr = bolConnErr
      WriteLog2 "連線失敗！", pDS
   End If
   
   If rsQuery.State = adStateOpen Then rsQuery.Close
   Set rsQuery = Nothing
   If cnnQurey.State = adStateOpen Then cnnQurey.Close
   Set cnnQurey = Nothing
   
End Function

Private Sub doConvert()
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim strFullFileName As String, strFileName As String
   Dim bolInTrans As Boolean
   'Add By Sindy 2019/11/26
   Dim bolGetPDF As Boolean, intRow As Integer
   Dim strCmd As String
   Dim process_id As Long
   Dim process_handle As Long
   Dim strMergeFN As String
   '2019/11/26 END
   
On Error GoTo ErrHnd
   
   Me.Enabled = False
   StatusBar1.Panels.Item(1).Text = "檢查中..."
   
   '釋放前次未轉完成定稿
   stSQL = "update letterprogress set lp08=null,lp13=null where lp08='" & pub_HostName & "' and lp09=0"
   cnnConnection.Execute stSQL, intR
   If intR > 0 Then
      lstHistory.AddItem Now & "--> 有 " & intR & " 筆未轉完成定稿已釋放!!"
   End If
            
   'Modify by Amy 2014/09/01 先轉客戶函再轉申請書 for P台灣案電子化
   'stSQL = "update letterprogress a set lp08='" & pub_HostName & "',lp13=sysdate where lp01 in (select min(lp01) from letterprogress b,letterdemand where b.lp09=0 and b.lp10='Y' and b.LP08 IS NULL and ld18(+)=lp01 and ld01 is not null)"
   'Modified by Morgan 2018/7/3 +回覆單的定稿也要轉(LP41='Y')
   'Modified by Morgan 2018/10/9 客戶函改判斷格式為橫式者 AND LD12='1'
   'Modified by Sindy 2019/10/29 + T客戶函 AND LD12='1' => AND ((LD05 in('P','PS','CFP','CPS') AND LD12='1') or (substr(LD05,1,1)='T' AND LD12 not in('5','6')))
   'Modify By Sindy 2019/12/5 AND ((LD05 in('P','PS','CFP','CPS') AND LD12='1') or (substr(LD05,1,1)='T' AND LD12 not in('5','6')))
   '                       => AND LD27 in('CUS','BLANK')
   'Modified by Morgan 2023/10/13 + and ld02||lpad(ld03,6,'0')<to_char(sysdate-1/24/60,'yyyymmddhh24miss') 延遲1分鐘轉PDF(T的收款寄證更新LP09會有時間差)
   stSQL = "update letterprogress a set lp08='" & pub_HostName & "',lp13=sysdate " & _
           "where lp01 in (select min(lp01) from letterprogress b,letterdemand " & _
                           "where b.lp09=0 and (b.lp10='Y' or b.lp41='Y') " & _
                           "and b.LP08 IS NULL and ld18(+)=lp01 and ld01 is not null " & _
                           "AND LD27 in('CUS','BLANK') and ld02||lpad(ld03,6,'0')<to_char(sysdate-1/24/60,'yyyymmddhh24miss')) "
   cnnConnection.Execute stSQL, intR
   If intR = 1 Then
      StatusBar1.Panels.Item(1).Text = "客戶函轉檔中..."
      
      'Modify by Amy 2014/09/01
      'stSQL = "select LP01,cp01,cp02,cp03,cp04,cp10,LD01,LD02,LD03 from letterprogress,caseprogress,LETTERDEMAND where lp09=0 and lp10='Y' AND lP08='" & pub_HostName & "' and cp09(+)=lP01 AND LD18(+)=CP09 and ld01 is not null ORDER BY LD02 DESC,LD03 DESC"
      'Modified by Morgan 2018/7/3 +回覆單的定稿也要轉(LP41='Y')
      'Modified by Sindy 2019/10/29 + T客戶函 AND LD12='1' => AND ((LD05 in('P','PS','CFP','CPS') AND LD12='1') or (substr(LD05,1,1)='T' AND LD12 not in('5','6')))
      'Modify By Sindy 2019/11/26 ORDER BY LD02 DESC,LD03 DESC => ORDER BY LD02 ASC,LD03 ASC
      'Modify By Sindy 2019/12/5 AND ((LD05 in('P','PS','CFP','CPS') AND LD12='1') or (substr(LD05,1,1)='T' AND LD12 not in('5','6')))
      '                       => AND LD27 in('CUS','BLANK')
      'Modify By Sindy 2021/1/12 + ,cp27
      stSQL = "select LP01,cp01,cp02,cp03,cp04,cp10,LD01,LD02,LD03,LD01,LP41,cp27 from letterprogress,caseprogress,LETTERDEMAND " & _
                   "where lp09=0 and (lp10='Y' or LP41='Y') AND lP08='" & pub_HostName & "' " & _
                   "and cp09(+)=lP01 AND LD18(+)=CP09 and ld01 is not null " & _
                   "AND LD27 in('CUS','BLANK') " & _
                   "ORDER BY LD02 ASC,LD03 ASC"
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         With rsQuery
            'Modified by Morgan 2015/1/26 本所號的追加聯合碼改各自判斷 Ex.P123456-1,P123456-0-01
            'Modified by Morgan 2018/7/3 +回覆單(.BLANK.PDF)
            'Modified by Morgan 2020/2/5 檔名改抓全部CP02
            'If .Fields("LP41") = "Y" Then
            '   strFileName = .Fields("cp01") & Val(.Fields("cp02")) & IIf(.Fields("cp04") <> "00", "-" & .Fields("cp03") & "-" & .Fields("cp04"), IIf(.Fields("cp03") <> "0", "-" & .Fields("cp03"), "")) & "." & .Fields("CP10") & ".BLANK.PDF"
            'Else
            '   strFileName = .Fields("cp01") & Val(.Fields("cp02")) & IIf(.Fields("cp04") <> "00", "-" & .Fields("cp03") & "-" & .Fields("cp04"), IIf(.Fields("cp03") <> "0", "-" & .Fields("cp03"), "")) & "." & .Fields("CP10") & ".CUS.PDF"
            'End If
            strFileName = PUB_CaseNo2FileName(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04")) & "." & .Fields("CP10")
            If .Fields("LP41") = "Y" Then
               strFileName = strFileName & ".BLANK.PDF"
            Else
               strFileName = strFileName & ".CUS.PDF"
            End If
            'end 2020/2/5
            
            'Added by Morgan 2015/11/17
            '檢查卷宗區檔案是否已存,若有則通知程序檢查
            stSQL = "update casepaperpdf set cpp05=cpp05 where cpp01='" & .Fields("LP01") & "' and upper(cpp02)=upper('" & strFileName & "')"
            cnnConnection.Execute stSQL, intR
            If intR <> 0 Then
               lstHistory.AddItem Now & "--> 檔案已存在 " & strFileName & "(" & .Fields("LP01") & ")"
               lstHistory.ListIndex = lstHistory.ListCount - 1
               
               'Modify By Sindy 2016/7/12 bolCaseDutyAgentMsg=False,才不會收件者休假彈出寄信給職代的訊息
               'PUB_SendMail "QPGMR", .Fields("LD01"), "", strFileName & "(" & .Fields("LP01") & ")卷宗區檔案已存在，客戶函轉檔取消！", "請確認卷宗區檔案正確性！若該檔非客戶函則請更名並利用定稿維護功能上傳。", , , , , , , "QPGMR"
               PUB_SendMail "QPGMR", .Fields("LD01"), "", strFileName & "(" & .Fields("LP01") & ")卷宗區檔案已存在，客戶函轉檔取消！", "請確認卷宗區檔案正確性！若該檔非客戶函則請更名並利用定稿維護功能上傳。", , , , , , , "QPGMR", , , , False
               
               stSQL = "update letterprogress set lp09=to_char(sysdate,'yyyymmdd') where lp01='" & .Fields("lp01") & "'"
               cnnConnection.Execute stSQL, intR
            Else
            'end 2015/11/17
               
               lstHistory.AddItem Now & "*** 客戶函 " & strFileName & " 轉檔開始 ***"
               lstHistory.ListIndex = lstHistory.ListCount - 1
               
               'Add By Sindy 2019/11/26
               '代表有需要合併的狀況
               If rsQuery.RecordCount > 1 Then
                  intRow = 0: strFullFileName = ""
                  strMergeFN = "" '組欲合併的檔案
                  rsQuery.MoveFirst
                  Do While Not rsQuery.EOF
                     intRow = intRow + 1
                     strFullFileName = "$$" & .Fields("LP01") & "." & Format(intRow, "000") & ".pdf"
                     bolGetPDF = ConvertLetter2PDF(.Fields("LD01"), .Fields("LD02"), .Fields("LD03"), strFullFileName)
                     If bolGetPDF = False Then
                        Exit Do
                     Else
                        strMergeFN = strMergeFN & IIf(strMergeFN <> "", " ", "") & ".\" & Mid(strFullFileName, InStrRev(strFullFileName, "\") + 1)
                     End If
                     rsQuery.MoveNext
                  Loop
                  '進行多檔合併
                  If bolGetPDF = True Then
                     rsQuery.MoveFirst
                     bolGetPDF = False '再預設為False,確認合併成功才會變成True
                     '切換至來源目錄
                     If App.path <> "." Then ChDir App.path
                     '合併
                     strFullFileName = "$$" & .Fields("LD01") & .Fields("LD02") & .Fields("LD03") & ".pdf"
                     strCmd = pub_PdftkEXE & " " & strMergeFN & " cat output .\" & strFullFileName
                     process_id = Shell(strCmd, vbHide)
                     process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
                     If process_handle <> 0 Then
                        For intI = 1 To 10
                           If PUB_CheckIsRunning(pub_PdftkName) = True Then
                              Sleep 1000
                           Else
                              Exit For
                           End If
                        Next
                        If intI > 10 Then
                           TerminateProcess process_handle, 0&
                           CloseHandle process_handle
                           lstHistory.AddItem Now & "--> 合併PDF失敗 " & strFileName
                           lstHistory.ListIndex = lstHistory.ListCount - 1
                           GoTo GetPDFErr
                        Else
                           CloseHandle process_handle
                        End If
                     Else
                        lstHistory.AddItem Now & "--> 合併PDF失敗 " & strFileName
                        lstHistory.ListIndex = lstHistory.ListCount - 1
                        GoTo GetPDFErr
                     End If
                     strFullFileName = App.path & "\" & strFullFileName
                     If Dir(strFullFileName) <> "" Then
                        bolGetPDF = True
                     End If
                  End If
               Else
                  bolGetPDF = ConvertLetter2PDF(.Fields("LD01"), .Fields("LD02"), .Fields("LD03"), strFullFileName)
               End If
               '2019/11/26 END
GetPDFErr:
               'Modified by Morgan 2017/9/15
               'If PUB_ConvLetter2PDF(.Fields("LD01"), .Fields("LD02"), .Fields("LD03"), strFullFileName) = True Then
               'Modify By Sindy 2019/11/26
               'If ConvertLetter2PDF(.Fields("LD01"), .Fields("LD02"), .Fields("LD03"), strFullFileName) = True Then
               If bolGetPDF = True Then
               '2019/11/26 END
                  Set oFile = oFileSys.GetFile(strFullFileName)
                  
                  'Added by Morgan 2015/11/6
                  cnnConnection.BeginTrans
                  bolInTrans = True
                  '再次鎖定(可能在轉定稿的過程中程序也在維護定稿)
                  stSQL = "Update letterprogress Set lp08=lp08 Where lp01='" & .Fields("lp01") & "' and lp08='" & pub_HostName & "'"
                  cnnConnection.Execute stSQL, intR
                  If intR = 1 Then
                  'end 2015/11/6
                   
      'Removed by Morgan 2015/11/17 不可刪,可能是上傳的
      '               PUB_DelFtpFile2 .Fields("LP01"), " and cpp02='" & strFileName & "'" 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
      '
      '               stSQL = "delete from CasePaperPDF where cpp01='" & .Fields("LP01") & "' and cpp02='" & strFileName & "'"
      '               cnnConnection.Execute stSQL, intR
      '               If intR = 1 Then
      '                  lstHistory.AddItem Now & "--> 卷宗區舊檔已刪除 " & strFileName
      '                  lstHistory.ListIndex = lstHistory.ListCount - 1
      '               End If
      'end 2015/11/17
                  
                     'Modify By Sindy 2015/5/18
                     'If SaveAttFile_PDF(.Fields("LP01"), strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, "4") = True Then
                     'Modified by Morgan 2015/11/17 pRaiseErr=True
                     If SaveAttFile_PDF(.Fields("LP01"), strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True) = True Then
                     '2015/5/18 END
                        stSQL = "update letterprogress set lp14=sysdate,lp09=to_char(sysdate,'yyyymmdd') where lp01='" & .Fields("lp01") & "'"
                        cnnConnection.Execute stSQL, intR
                        lstHistory.AddItem Now & "--> 已轉入卷宗區 " & strFileName
                        lstHistory.ListIndex = lstHistory.ListCount - 1
                        If .Fields("cp27") <> "" Then PUB_UpdateLP03 .Fields("LP01") 'Add By Sindy 2021/1/12
                     Else
                        stSQL = "update letterprogress set lp08=null,lp13=null where lp01='" & .Fields("lp01") & "'"
                        cnnConnection.Execute stSQL, intR
                        lstHistory.AddItem Now & "--> 轉入卷宗區失敗 " & strFileName
                        lstHistory.ListIndex = lstHistory.ListCount - 1
                     End If
                     
                  'Added by Morgan 2015/11/6
                  Else
                     lstHistory.AddItem Now & "--> 轉入取消 " & strFileName
                     lstHistory.ListIndex = lstHistory.ListCount - 1
                  End If
                  cnnConnection.CommitTrans
                  bolInTrans = False
                  'end 2015/11/6
                  
                  Kill strFullFileName
               Else
                  stSQL = "update letterprogress set lp08=null,lp13=null where lp01='" & .Fields("lp01") & "'"
                  cnnConnection.Execute stSQL, intR
                  lstHistory.AddItem Now & "--> pdf產生失敗 " & strFileName
                  lstHistory.ListIndex = lstHistory.ListCount - 1
               End If
            End If 'Added by Morgan 2015/11/17
         End With
      End If
   Else
      'Modify by Amy 2014/09/01 +轉申請書
'      If P台灣案電子化啟用日 <= Val(strSrvDate(1)) Then 'Modify by Amy 2014/11/14拿掉日期上線試run
        '釋放前次未轉完成申請書
        stSQL = "Update AppForm Set AF02=null,Af04=null Where AF02='" & pub_HostName & "' and AF03=0"
        cnnConnection.Execute stSQL, intR
        If intR > 0 Then
            lstHistory.AddItem Now & "--> 有 " & intR & " 筆未轉完成定稿已釋放!!"
        End If
   
        'Modified by Morgan 2015/10/22 +大陸指示信
        'Modify By Sindy 2019/12/5 AND LD12 in ('5','6')
        '                       => AND LD27 ='DATA'
        'Modified by Morgan 2020/6/29 排除CFP案
        stSQL = "Update AppForm a Set AF02='" & pub_HostName & "',AF04=sysdate " & _
                    "Where AF01 in (Select min(AF01) From AppForm b,LetterDemand " & _
                    "Where b.AF03=0 And b.AF02 Is Null And LD18(+)=AF01 And LD01 Is not Null " & _
                    "AND LD27 ='DATA'" & _
                    ") and not exists(select * from caseprogress where cp09=af01 and cp01='CFP')"
        cnnConnection.Execute stSQL, intR
        If intR = 1 Then
            StatusBar1.Panels.Item(1).Text = "申請書/指示信轉檔中..."
            
            'Modify By Sindy 2019/12/5 AND LD12 in ('5','6')
            '                       => AND LD27 ='DATA'
            stSQL = "Select AF01,cp01,cp02,cp03,cp04,cp10,LD01,LD02,LD03 From AppForm,caseprogress,LetterDemand " & _
                        "Where AF03=0 And AF02='" & pub_HostName & "' And cp09(+)=AF01 And LD18(+)=CP09 " & _
                        "And LD01 Is not Null AND LD27 ='DATA'" & _
                        "Order by LD02 Desc,LD03 Desc"
            Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
            If intR = 1 Then
                With rsQuery
             
               'Modified by Morgan 2015/1/26 本所號的追加聯合碼改各自判斷 Ex.P123456-1,P123456-0-01
                'Modified by Morgan 2020/2/5 檔名改抓全部CP02
                'strFileName = .Fields("cp01") & Val(.Fields("cp02")) & IIf(.Fields("cp04") <> "00", "-" & .Fields("cp03") & "-" & .Fields("cp04"), IIf(.Fields("cp03") <> "0", "-" & .Fields("cp03"), "")) & "." & .Fields("CP10") & ".DATA.PDF"
                strFileName = PUB_CaseNo2FileName(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04")) & "." & .Fields("CP10") & ".DATA.PDF"
                'end 2020/2/5
                
               'Added by Morgan 2015/11/17
               '檢查卷宗區檔案是否已存,若有則通知程序檢查
               stSQL = "update casepaperpdf set cpp05=cpp05 where cpp01='" & .Fields("AF01") & "' and upper(cpp02)=upper('" & strFileName & "')"
               cnnConnection.Execute stSQL, intR
               If intR <> 0 Then
                  lstHistory.AddItem Now & "--> 檔案已存在 " & strFileName & "(" & .Fields("AF01") & ")"
                  lstHistory.ListIndex = lstHistory.ListCount - 1
                  
                  'Modify By Sindy 2016/7/12 bolCaseDutyAgentMsg=False,才不會收件者休假彈出寄信給職代的訊息
                  'PUB_SendMail "QPGMR", .Fields("LD01"), "", strFileName & "(" & .Fields("AF01") & ")卷宗區檔案已存在，申請書/指示信轉檔取消！", "請確認卷宗區檔案正確性！若該檔非申請書/指示信則請更名並利用定稿維護功能上傳。", , , , , , , "QPGMR"
                  PUB_SendMail "QPGMR", .Fields("LD01"), "", strFileName & "(" & .Fields("AF01") & ")卷宗區檔案已存在，申請書/指示信轉檔取消！", "請確認卷宗區檔案正確性！若該檔非申請書/指示信則請更名並利用定稿維護功能上傳。", , , , , , , "QPGMR", , , , False
                  
                  stSQL = "Update AppForm Set AF03=to_char(sysdate,'yyyymmdd') Where AF01='" & .Fields("AF01") & "'"
                  cnnConnection.Execute stSQL, intR
               Else
               'end 2015/11/17
                
               lstHistory.AddItem Now & "*** 申請書/指示信 " & strFileName & " 轉檔開始 ***"
               lstHistory.ListIndex = lstHistory.ListCount - 1
                'Modified by Morgan 2017/9/15
                'If PUB_ConvLetter2PDF(.Fields("LD01"), .Fields("LD02"), .Fields("LD03"), strFullFileName) = True Then
                If ConvertLetter2PDF(.Fields("LD01"), .Fields("LD02"), .Fields("LD03"), strFullFileName) = True Then
                  Set oFile = oFileSys.GetFile(strFullFileName)
                  
                  'Added by Morgan 2015/11/6
                  cnnConnection.BeginTrans
                  bolInTrans = True
                  '再次鎖定(可能在轉定稿的過程中程序也在維護定稿)
                  stSQL = "Update AppForm Set AF05=AF05 Where AF01='" & .Fields("AF01") & "' and AF02='" & pub_HostName & "'"
                  cnnConnection.Execute stSQL, intR
                  If intR = 1 Then
                  'end 2015/11/6
                  
'Removed by Morgan 2015/11/17 不可刪,可能是上傳的
'                    PUB_DelFtpFile2 .Fields("AF01"), " and cpp02='" & strFileName & "'" 'Added by Morgan 2015/4/15 檔案改放 FTP,必須在DB資料刪除前執行
'
'                    stSQL = "Delete From CasePaperPDF Where cpp01='" & .Fields("AF01") & "' And cpp02='" & strFileName & "'"
'                    cnnConnection.Execute stSQL, intR
'                    If intR = 1 Then
'                       lstHistory.AddItem Now & "--> 卷宗區舊檔已刪除 " & strFileName
'                       lstHistory.ListIndex = lstHistory.ListCount - 1
'                    End If
'end 2015/11/17
                      
                    'Modify By Sindy 2015/5/18
                    'If SaveAttFile_PDF(.Fields("AF01"), strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, "4") = True Then
                    'Modified by Morgan 2015/11/17 pRaiseErr=True
                    If SaveAttFile_PDF(.Fields("AF01"), strFullFileName, strFileName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True) = True Then
                    '2015/5/18 END
                        stSQL = "Update AppForm Set AF05=sysdate,AF03=to_char(sysdate,'yyyymmdd') Where AF01='" & .Fields("AF01") & "'"
                        cnnConnection.Execute stSQL, intR
                        lstHistory.AddItem Now & "--> 已轉入卷宗區 " & strFileName
                        lstHistory.ListIndex = lstHistory.ListCount - 1
                    Else
                       stSQL = "Update AppForm Set AF02=null,AF04=null Where AF01='" & .Fields("AF01") & "'"
                        cnnConnection.Execute stSQL, intR
                        lstHistory.AddItem Now & "--> 轉入卷宗區失敗 " & strFileName
                        lstHistory.ListIndex = lstHistory.ListCount - 1
                    End If
                    
                  'Added by Morgan 2015/11/6
                  Else
                     lstHistory.AddItem Now & "--> 轉入取消 " & strFileName
                     lstHistory.ListIndex = lstHistory.ListCount - 1
                  End If
                  cnnConnection.CommitTrans
                  bolInTrans = False
                  'end 2015/11/6
                  
                  Kill strFullFileName
                  
                Else
                  stSQL = "Update AppForm Set AF02=null,AF04=null Where AF01='" & .Fields("AF01") & "'"
                  cnnConnection.Execute stSQL, intR
                  lstHistory.AddItem Now & "--> pdf產生失敗 " & strFileName
                  lstHistory.ListIndex = lstHistory.ListCount - 1
                End If
                'lstHistory.AddItem Now & "*** 申請書/指示信轉檔結束 ***"
                'lstHistory.ListIndex = lstHistory.ListCount - 1
                
                End If 'Added by Morgan 2015/11/17
                
                End With
            End If
        Else
            tmrPolling.Interval = 6000
        End If
'      Else
'        'P台灣案電子化未上線前使用
'        tmrPolling.Interval = 6000
'        'lstHistory.AddItem Now & "--> 無待轉定稿"
'        'lstHistory.ListIndex = lstHistory.ListCount - 1
'      End If
      'end 2014/09/01
      'end 2014/11/14拿掉日期上線試run
   End If
   
   StatusBar1.Panels.Item(1).Text = "等待中..."
   
ErrHnd:
   If bolInTrans Then cnnConnection.RollbackTrans 'Added by Morgan 2015/11/6

   If Err.Number <> 0 Then
      lstHistory.AddItem Now & "-->" & Err.Description & "-" & Err.Number
   End If
   
   Set rsQuery = Nothing
   Me.Enabled = True
End Sub

Private Sub tmrPolling_Timer()
   Static strDate As String
   Static iTime As Integer
   Static bolReConnect As Boolean
   Static bolReConnectClose As Boolean
    
    
   'Added by Morgan 2015/8/25
   '員工檔改結構需重新連線,否則會發生Trigger錯誤
   If strDate = "" Then strDate = Format(Now, "YYYYMMDD")
   If strDate <> Format(Now, "YYYYMMDD") And Format(Now, "HH") > 1 Then
      strDate = Format(Now, "YYYYMMDD")
      bolReConnect = True
      bolReConnectClose = True
      tmrPolling.Interval = 60000
   End If
      
   If bolReConnect Then
      If bolReConnectClose Then
         If cnnConnection.State = adStateOpen Then
            lstHistory.AddItem Now & "--> 結束連線"
            WriteLog True
            cnnConnection.Close
            lstHistory.AddItem Now & "--> 連線已結束"
            WriteLog True
            
            KillFile 'Add By Sindy 2019/11/18
         End If
         bolReConnectClose = False
      Else
         If cnnConnection.State = adStateClosed Then
            lstHistory.AddItem Now & "--> 重新連線"
            WriteLog True
            '重新連線
            Me.Caption = Me.Tag
            If fConnect(Me) = False Then
               lstHistory.AddItem Now & "--> 重新連線失敗"
               WriteLog True
               Exit Sub
            Else
               lstHistory.AddItem Now & "--> 重新連線成功"
               WriteLog True
               tmrPolling.Interval = 1000
            End If
         End If
         bolReConnect = False
      End If
   Else
      doConvert
   End If
   'end 2015/8/25
   
   iTime = iTime + 1
   If lstHistory.ListCount > 120 Then
      WriteLog
      iTime = 0
   ElseIf iTime > 60 Then
      If lstHistory.ListCount > 1 Then
         WriteLog True
      End If
      iTime = 0
   End If
   
End Sub

Private Sub WriteLog(Optional pAll As Boolean)
   Dim stLogFolder As String, stLogFile As String, ffa As Integer
   Dim ii As Integer, iMax As Integer
   
On Error GoTo ErrHnd
   
   If pAll = True Then
      iMax = lstHistory.ListCount
   Else
      iMax = 100
   End If
      
   stLogFolder = App.path & "\" & App.EXEName & "Log"
   If Dir(stLogFolder, vbDirectory) = "" Then
      MkDir stLogFolder
   End If
   
   'log保留一年(清除前一年的log)
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww") - 100) & ".log"
   If Dir(stLogFile) <> "" Then
      Kill stLogFile
   End If
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww")) & ".log"
   
   ffa = FreeFile
On Error GoTo ErrHnd2
   Open stLogFile For Append As ffa
   For ii = 1 To iMax
      If lstHistory.ListCount = 0 Then Exit For
      Print #ffa, lstHistory.List(0)
      lstHistory.RemoveItem 0
   Next
   
ErrHnd2:
   Close ffa
   lstHistory.AddItem Now & "--> 較早訊息移入LOG檔(" & ii - 1 & ")"
   
ErrHnd:

End Sub
'Added by Morgan 2017/1/11
Private Function SaveSQL2File(pSQL As String, pFilePath As String) As Boolean
   Dim stFileName As String, ffa As Integer
   Dim ii As Integer, iMax As Integer
   
On Error GoTo ErrHnd
      
   stFileName = App.path & "\ActiveSQL.txt"
   
   If Dir(stFileName) <> "" Then Kill stFileName
   
   ffa = FreeFile
   Open stFileName For Append As ffa
   Print #ffa, pSQL
   Close ffa
   
   pFilePath = stFileName
   SaveSQL2File = True
   
ErrHnd:
   If ffa <> 0 Then Close ffa
End Function

Private Sub WriteLog2(pLog As String, pDS As String)
   Dim stLogFolder As String, stLogFile As String, ffa As Integer
   
On Error GoTo ErrHnd
      
   stLogFolder = App.path & "\" & App.EXEName & "Log"
   If Dir(stLogFolder, vbDirectory) = "" Then
      MkDir stLogFolder
   End If
   
   'log保留一年(清除前一年的log)
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww") - 100) & pDS & ".log"
   If Dir(stLogFile) <> "" Then
      Kill stLogFile
   End If
   stLogFile = stLogFolder & "\" & (Format(Now, "yyyyww")) & pDS & ".log"
   
   ffa = FreeFile

   Open stLogFile For Append As ffa
   Print #ffa, Now & "=>" & pLog
   
ErrHnd:

On Error Resume Next
   If ffa > 0 Then Close ffa

End Sub

Private Function ConvertLetter2PDF(pLD01 As String, pLD02 As String, pLD03 As String, Optional ByRef pFileName As String) As Boolean
   If pub_Word2Pdf Then
      ConvertLetter2PDF = PUB_ConvLetter2PDFbyWord(pLD01, pLD02, pLD03, pFileName)
   Else
      ConvertLetter2PDF = PUB_ConvLetter2PDF(pLD01, pLD02, pLD03, pFileName)
   End If
End Function

'Added by Lydia 2017/12/12 FCP案件命名通知信-郵件主旨和收件人
'Modified by Lydia 2018/03/07 +副本iCC
Private Sub GetMailSub(ByVal iKind As String, ByRef iSub As String, ByRef iMan As String, _
                                   ByVal cp01 As String, ByVal cp02 As String, ByVal cp03 As String, ByVal cp04 As String, ByVal cp06 As String, ByVal cp07 As String, _
                                   ByVal pDate As String, ByVal pTime As String, ByVal pT04 As String, ByVal pT10 As String, ByRef iCC As String)
Dim Str01 As String
Dim intR As Integer
Dim rsB1 As New ADODB.Recordset
Dim strSpec As String 'Added by Lydia 2024/11/15
     
     'Added by Lydia 2024/11/15
     If Asc(Left(iKind, 1)) < 48 Or Asc(Left(iKind, 1)) > 57 Then
        'A: 只發工程師,不發程序
        strSpec = Left(iKind, 1)
        iKind = Mid(iKind, 2)
     End If
     'end 2024/11/15
     
     '郵件主旨
     'Modified by Lydia 2019/01/02 逾期後,增加8:00和13:30的逾期通知信
     'If iKind = "2" Then
     If Val(iKind) > 2 Then
         Str01 = "已逾期通知(" & Val(iKind) - 1 & ")："
     ElseIf iKind = "2" Then
     'end 2019/01/02
         Str01 = "已逾期通知："
     Else
         Str01 = "到期前通知："
     End If
     iSub = Str01 & IIf(pDate <> "", "急件！", "") & cp01 & "-" & cp02 & IIf(cp03 & cp04 <> "000", "-" & cp03 & "-" & cp04, "") & "新案命名"
     'Modified by Lydia 2018/12/17 去掉完成期限(所限)
     'If pDate <> "" Or cp06 <> "" Then
     If pDate <> "" Then
          iSub = iSub & "("
          'Remove by Lydia 2018/12/17 將完成期限刪除以免工程師誤判命名deadline(ex.FCP-060061工程師堅持要到所限才完成)
          'If cp06 <> "" Then
          '   iSub = iSub & "完成期限：" & ChangeTStringToTDateString(TransDate(cp06, 1)) & IIf(InStr(iSub, "急件") > 0, "，", "")
          'End If
          'end 2018/12/17
          If InStr(iSub, "急件") > 0 Then
             iSub = iSub & "譯畢期限：" & ChangeTStringToTDateString(TransDate(pDate, 1)) & " " & Format(pTime, "00:00")
          End If
          iSub = iSub & ")"
     End If
       
     '收件人
     'Modified by Lydia 2018/03/06 "," =>";"
     'Modified by Lydia 2018/03/27
     'iMan = pT04 & ";" & pT10 & ";"
     iMan = pT04 & ";" & IIf(pT10 <> "", pT10 & ";", "")
     iCC = "" 'Added by Lydia 2018/03/07
     'Modified by Lydia 2019/01/02 逾期後,增加8:00和13:30的逾期通知信
     'If iKind = "2" Then  '已逾期+ FCP程序管制人和主管
     'Modified by Lydia 2024/11/15 A: 只發工程師,不發程序
     If Val(iKind) >= 2 And strSpec <> "A" Then
         Str01 = PUB_GetFCPHandler(cp01, cp02, cp03, cp04)
         If Str01 <> "" Then
              iMan = iMan & Str01 & ";"
              'Modified by Lydia 2018/03/07 st52=>nvl(st52,'N')
              Str01 = "select nvl(st52,'N') from staff where st01='" & Str01 & "' "
              Set rsB1 = ClsLawReadRstMsg(intR, Str01)
              If intR = 1 Then
                   'Modified by Lydia 2018/03/07 改副本
                   'iMan = iMan & rsB1.Fields(0) & ";"
                   If "" & rsB1.Fields(0) <> "N" Then iCC = iCC & rsB1.Fields(0) & ";"
               End If
         End If
     End If
     
     Set rsB1 = Nothing
End Sub

'Added by Lydia 2017/12/12 FCP案件命名通知信
Private Sub tmrMail_Timer()
Dim strTime As String
Dim intM As Integer
Dim rsRd As New ADODB.Recordset
Dim strSub As String, strTo As String
Dim strDate As String
Dim strCC As String 'Added by Lydia 2018/03/07 副本
Dim strAct As String 'Added by Morgan 2019/7/9
Dim strNoList As String 'Added by Lydia 2024/11/15

On Error GoTo ErrFCPmail 'Added by Lydia 2018/03/08
   
   'Add by Amy 2019/07/11 7點-19點後 才執行
   If Not (Val(Format(Now, "hhmm")) >= Val(strTimeS) And Val(Format(Now, "hhmm")) <= Val(strTimeE)) Then Exit Sub

   'Added by Lydia 2018/03/08 暫存X分鐘
   tmrMail.Tag = Val(tmrMail.Tag) + 1
   If Val(tmrMail.Tag) < 5 Then Exit Sub
   tmrMail.Tag = 0
   'end 2018/03/08
   
   'Added by Morgan 2023/8/18
   strAct = "檢查待執行紀錄 (Morgan)"
   If cnnConnection.State <> adStateOpen Then
      lstHistory.AddItem Now & "--> 中斷連線 " & strAct
      WriteLog True
      Exit Sub
   End If
   lstHistory.AddItem Now & "--> 開始 " & strAct
   PUB_ChkWaitExeRec
   lstHistory.AddItem Now & "--> 結束 " & strAct
   'end 2023/8/18
   
   
   'Add by Morgan 2019/6/28
   'Modified by Morgan 2019/7/8 +True
   'Modified by Morgan 2019/7/9 +log
   'Modified by Morgan 2019/7/11 +連線判斷
   strAct = "寄發 QPGMR 郵件暫存 (Morgan)"
   If cnnConnection.State <> adStateOpen Then
      lstHistory.AddItem Now & "--> 中斷連線 " & strAct
      WriteLog True
      Exit Sub
   End If
   lstHistory.AddItem Now & "--> 開始 " & strAct
   PUB_SendMailCache True
   lstHistory.AddItem Now & "--> 結束 " & strAct
   strAct = "FCP案件命名通知信"
   'end 2019/6/28
   
 'Modified by Lydia 2018/03/07 因為12點換日的時候,常駐程式不一定會讀取到變數,改用machine time
'If strSrvDate(1) <= FCP案件命名啟用日 Then Exit Sub

    'strDate = strSrvDate(1)
    'strTime = Left(Format(ServerTime, "000000"), 4)
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhmm")
    If strSrvDate(1) <= FCP案件命名啟用日 Then Exit Sub
    'end 2018/03/07
    
    'Added by Lydia 2018/03/08 判斷有連線才做
    If cnnConnection.State <> adStateOpen Then
         lstHistory.AddItem Now & "--> 中斷連線 FCP案件命名通知信"
         WriteLog True
         Exit Sub
    End If
    lstHistory.AddItem Now & "--> 開始 FCP案件命名通知信"
    'end 2018/03/08
    
    'Addded by Lydia 2019/01/02
    '>2 將命名作業的"逾期通知"改成上午(8:00)和下午(13:30)各通知一次，直至工程師命名為止by Phoebe
    'Modified by Lydia 2025/06/02 因為執行時間會有非8點的情況,改用記錄判斷; ex.FCP-073588自5/23逾期後,只有5/30和5/31通知
    'If strTime = "0800" Or strTime = "1330" Then
    strExc(1) = ""
    If Val(strTime) > 800 Then
       strExc(2) = IIf(Val(strTime) >= 1330, "2", "1")
       strExc(0) = "select * from addressa4list where aal01='TCT已逾期' and aal02=to_char(sysdate,'yyyymmdd') and aal03='" & strExc(2) & "' "
       intM = 1
       Set rsRd = ClsLawReadRstMsg(intM, strExc(0))
       If intM = 0 Then
         strSql = "Delete from addressa4list where aal01='TCT已逾期' and aal02<to_char(sysdate,'yyyymmdd') "
         cnnConnection.Execute strSql
         strSql = "Insert into addressa4list (aal01,aal02,aal03,aal04) values ('TCT已逾期',to_char(sysdate,'yyyymmdd'),'" & strExc(2) & "','QPGMR') "
         cnnConnection.Execute strSql
         strExc(1) = "Y"
       End If
    End If
    If strExc(1) = "Y" Then
    'end 2025/06/02
         'Modified by Lydia 2025/01/17 暫不認領TCN16=Y的取消日期記錄在命名記錄TransCaseTitle.TCT121,TCT122
         'strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,CP66)) exp_date,NVL(TCT03,CP67) exp_time, " & _
                     "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07,TCT115 " & _
                     "FROM TransCaseTitle,CASEPROGRESS " & _
                     "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') >= '2' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                     "AND CP01='FCP' AND CP159=0 AND (NVL(TCT02,WORKDAYADD(2,CP66))<" & strDate & " OR (NVL(TCT02,WORKDAYADD(2,CP66))=" & strDate & " AND NVL(TCT03,CP67) <=" & Val(strTime) & " )) "
         strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66))) exp_date,NVL(TCT03,nvl(tct121,CP67)) exp_time, " & _
                     "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07,TCT115 " & _
                     "FROM TransCaseTitle,CASEPROGRESS " & _
                     "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') >= '2' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                     "AND CP01='FCP' AND CP159=0 AND (NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66)))<" & strDate & " OR (NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66)))=" & strDate & " AND NVL(TCT03,nvl(tct121,CP67)) <=" & Val(strTime) & " )) "
         strExc(1) = Replace(Replace(UCase(strSql), "WORKDAYADD(2,", "WORKDAYADD(3,"), "='FCP'", "<>'FCP'") 'FMP案的期限多+1天
         strSql = strSql & " UNION ALL " & strExc(1) & " ORDER BY 2,3"
         intM = 1
         Set rsRd = ClsLawReadRstMsg(intM, strSql)
         If intM = 1 Then
             With rsRd
                  .MoveFirst
                  Do While Not .EOF
                       strSub = "": strTo = "": strCC = ""
                       Call GetMailSub(Val("" & rsRd.Fields("TCT115")) + 1, strSub, strTo, .Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), "" & .Fields("cp06"), "" & .Fields("cp07"), "" & .Fields("TCT02"), "" & .Fields("TCT03"), "" & .Fields("TCT04"), "" & .Fields("TCT10"), strCC)
                       If strTo <> "" Then
                            If strTo <> "" Then
                                strTo = Mid(strTo, 1, Len(strTo) - 1)
                            End If
                            If strCC <> "" Then
                                strCC = Mid(strCC, 1, Len(strCC) - 1)
                            End If
                            If "" & .Fields("TCT07") <> "" And "" & .Fields("TCT10") = "" Then
                                 strTo = strTo & ";" & .Fields("TCT07")
                            End If

                            PUB_SendMail "QPGMR", strTo, "", strSub, "同主旨", , , , , , strCC
                            strSql = "Update TransCaseTitle set TCT115=" & CNULL(Val("" & rsRd.Fields("TCT115")) + 1) & "  Where TCT01='" & .Fields("TCT01") & "' "
                            cnnConnection.Execute strSql, intM

                            lstHistory.AddItem Now & "--> " & strSub & " 收件者: " & strTo & IIf(strCC <> "", " 副本:" & strCC, "")
                            strNoList = strNoList & IIf(strNoList <> "", ",", "") & strNoList 'Added by Lydia 2024/11/15
                       End If
                       .MoveNext
                  Loop
             End With
         End If
          
         'Added by Lydia 2024/11/15 針對非英說案且有設定待英文本翻譯之重新命名作業，增加逾期通知Email作業。
         strSql = " SELECT TCT01,WORKDAYADD(2,TCN26) AS EXP_DATE,'0800' AS EXP_TIME,TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,NVL(TCT115,1) AS TCT115 " & _
                  " From TRANSCASETITLE, CASEPROGRESS, TRACKINGCASENAME WHERE TCT01=CP09(+) AND CP159=0 AND TCT01=TCN05(+) AND NVL(TCT04,'N')<>'N'" & _
                  " AND NVL(TCT05,0)= 0 AND TCN13='3' AND WORKDAYADD(2,TCN26)<=" & strSrvDate(1)
         If strNoList <> "" Then
            strSql = strSql & " AND TCT01 NOT IN (" & GetAddStr(strNoList) & ")"
         End If
         strSql = strSql & " ORDER BY 2,1 "
         intM = 1
         Set rsRd = ClsLawReadRstMsg(intM, strSql)
         If intM = 1 Then
            With rsRd
               .MoveFirst
               Do While Not .EOF
                  strSub = "": strTo = "": strCC = ""
                  'A: 只發工程師,不發程序
                  Call GetMailSub("A" & Val("" & rsRd.Fields("TCT115")) + 1, strSub, strTo, .Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), "" & .Fields("exp_date"), "" & .Fields("exp_date"), "" & .Fields("TCT02"), "" & .Fields("TCT03"), "" & .Fields("TCT04"), "" & .Fields("TCT10"), strCC)
                  If strTo <> "" Then
                     If strTo <> "" Then
                         strTo = Mid(strTo, 1, Len(strTo) - 1)
                     End If
                     If strCC <> "" Then
                         strCC = Mid(strCC, 1, Len(strCC) - 1)
                     End If
                     If "" & .Fields("TCT07") <> "" And "" & .Fields("TCT10") = "" Then
                          strTo = strTo & ";" & .Fields("TCT07")
                     End If
                     'Modified by Lydia 2024/11/22 主旨+〔非英說案: 已收參考本〕
                     PUB_SendMail "QPGMR", strTo, "", strSub & "〔非英說案: 已收參考本〕", "同主旨", , , , , , strCC
                     strSql = "Update TransCaseTitle set TCT115=" & CNULL(Val("" & rsRd.Fields("TCT115")) + 1) & "  Where TCT01='" & .Fields("TCT01") & "' "
                     cnnConnection.Execute strSql, intM

                     lstHistory.AddItem Now & "--> " & strSub & " 收件者: " & strTo & IIf(strCC <> "", " 副本:" & strCC, "")
                  End If
                  .MoveNext
               Loop
            End With
         End If
         'end 2024/11/15
    End If
    'end 2019/01/02
    
    '2-已逾期通知
    'Modified by Lydia 2018/03/27 與Jack討論未分命名人員前,要發逾期通知信; 分命名人員後(TCT115=Null),再發一次通知信
    'strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,CP66)) exp_date,NVL(TCT03,CP67) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') < '2' AND NVL(TCT10,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND (NVL(TCT02,WORKDAYADD(2,CP66))<" & strDate & " OR (NVL(TCT02,WORKDAYADD(2,CP66))=" & strDate & " AND NVL(TCT03,CP67) <=" & Val(strTime) & " )) " & _
                " order by 2,3 "
    'Modified by Lydia 2018/06/05 FMP案因為從樓上下來,所以期限多+1天
    'strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,CP66)) exp_date,NVL(TCT03,CP67) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') < '2' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND (NVL(TCT02,WORKDAYADD(2,CP66))<" & strDate & " OR (NVL(TCT02,WORKDAYADD(2,CP66))=" & strDate & " AND NVL(TCT03,CP67) <=" & Val(strTime) & " )) " & _
                " order by 2,3 "
    'Modified by Lydia 2018/08/27 判斷取消收文不通知(ex.FCP-59484)
    'Modified by Lydia 2025/01/17 暫不認領TCN16=Y的取消日期記錄在命名記錄TransCaseTitle.TCT121,TCT122
    'strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,CP66)) exp_date,NVL(TCT03,CP67) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') < '2' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND CP01='FCP' AND CP159=0 AND (NVL(TCT02,WORKDAYADD(2,CP66))<" & strDate & " OR (NVL(TCT02,WORKDAYADD(2,CP66))=" & strDate & " AND NVL(TCT03,CP67) <=" & Val(strTime) & " )) "
    strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66))) exp_date,NVL(TCT03,nvl(tct121,CP67)) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') < '2' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND CP01='FCP' AND CP159=0 AND (NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66)))<" & strDate & " OR (NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66)))=" & strDate & " AND NVL(TCT03,nvl(tct121,CP67)) <=" & Val(strTime) & " )) "
    strExc(1) = Replace(Replace(UCase(strSql), "WORKDAYADD(2,", "WORKDAYADD(3,"), "='FCP'", "<>'FCP'") 'FMP案的期限多+1天
    strSql = strSql & " UNION ALL " & strExc(1) & " ORDER BY 2,3"
    'end 2018/06/05
    intM = 1
    Set rsRd = ClsLawReadRstMsg(intM, strSql)
    If intM = 1 Then
        With rsRd
             .MoveFirst
             Do While Not .EOF
                  'Modified by Lydia 2018/03/07 + strCC
                   strSub = "": strTo = "": strCC = ""
                   Call GetMailSub("2", strSub, strTo, .Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), "" & .Fields("cp06"), "" & .Fields("cp07"), "" & .Fields("TCT02"), "" & .Fields("TCT03"), "" & .Fields("TCT04"), "" & .Fields("TCT10"), strCC)
                   'end 2018/03/07
                   If strTo <> "" Then
                       'Added by Lydia 2018/03/07
                        If strTo <> "" Then
                            strTo = Mid(strTo, 1, Len(strTo) - 1)
                        End If
                        If strCC <> "" Then
                            strCC = Mid(strCC, 1, Len(strCC) - 1)
                        End If
                       'end 2018/03/07
                       'Added by Lydia 2018/10/31 若停在分案主任一併通知 (ex.FCP-59822)
                       If "" & .Fields("TCT07") <> "" And "" & .Fields("TCT10") = "" Then
                             strTo = strTo & ";" & .Fields("TCT07")
                       End If
                       'end 2018/10/31
                       
                       'Modifed by Lydia 2018/03/07 + strCC
                       PUB_SendMail "QPGMR", strTo, "", strSub, "同主旨", , , , , , strCC
                       strSql = "update TransCaseTitle set TCT115='2'  Where TCT01='" & .Fields("TCT01") & "' "
                       cnnConnection.Execute strSql, intM
                       'Added by Lydia 2018/03/08 記錄log
                       lstHistory.AddItem Now & "--> " & strSub & " 收件者: " & strTo & IIf(strCC <> "", " 副本:" & strCC, "")
                   End If
                   .MoveNext
             Loop
        End With
    End If
    
    '到期前1小時通知
    'Modified by Lydia 2018/03/19 凌晨(ex.0001)用#格式,於DateAdd會出錯
    'strTime = Left(Format(DateAdd("n", 60, Format(strTime & "00", "##:##:##")), "HHMMSS"), 4)
    strTime = Left(Format(DateAdd("n", 60, Format(strTime & "00", "00:00:00")), "HHMMSS"), 4)
    '1-到期前通知
    'Modified by Lydia 2018/03/27 與Jack討論未分命名人員前,要發逾期通知信; 分命名人員後(TCT115=Null),再發一次通知信
    'strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,CP66)) exp_date,NVL(TCT03,CP67) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') = '0' AND NVL(TCT10,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND NVL(TCT02,WORKDAYADD(2,CP66))=" & strDate & " AND NVL(TCT03,CP67) <=" & Val(strTime) & _
                " order by 2,3 "
    'Modified by Lydia 2018/06/05 FMP案因為從樓上下來,所以期限多+1天
    'strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,CP66)) exp_date,NVL(TCT03,CP67) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') = '0' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND NVL(TCT02,WORKDAYADD(2,CP66))=" & strDate & " AND NVL(TCT03,CP67) <=" & Val(strTime) & _
                " order by 2,3 "
    'Modified by Lydia 2018/08/27 判斷取消收文不通知(ex.FCP-59484)
    'Modified by Lydia 2025/01/17 暫不認領TCN16=Y的取消日期記錄在命名記錄TransCaseTitle.TCT121,TCT122;
    'strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,CP66)) exp_date,NVL(TCT03,CP67) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') = '0' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND CP01='FCP' AND CP159=0 AND NVL(TCT02,WORKDAYADD(2,CP66))=" & strDate & " AND NVL(TCT03,CP67) <=" & Val(strTime)
    strSql = "SELECT TCT01,NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66))) exp_date, NVL(TCT03,nvl(tct121,CP67)) exp_time, " & _
                "TCT02,TCT03,TCT04,TCT07,TCT10,CP01,CP02,CP03,CP04,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS " & _
                "WHERE TCT01=CP09(+) AND CP158=0 AND NVL(TCT115,'0') = '0' AND NVL(TCT04,'N')<>'N' AND NVL(TCT05,0)= 0 " & _
                "AND CP01='FCP' AND CP159=0 AND NVL(TCT02,WORKDAYADD(2,nvl(tct121,CP66)))=" & strDate & " AND NVL(TCT03,nvl(tct121,CP67)) <=" & Val(strTime)
    strExc(1) = Replace(Replace(UCase(strSql), "WORKDAYADD(2,", "WORKDAYADD(3,"), "='FCP'", "<>'FCP'") 'FMP案的期限多+1天
    strSql = strSql & " UNION ALL " & strExc(1) & " ORDER BY 2,3"
    'end 2018/06/05
    intM = 1
    Set rsRd = ClsLawReadRstMsg(intM, strSql)
    If intM = 1 Then
        With rsRd
             .MoveFirst
             Do While Not .EOF
                   'Modified by Lydia 2018/03/07 + strCC
                   strSub = "": strTo = ""
                   Call GetMailSub("1", strSub, strTo, .Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"), "" & .Fields("cp06"), "" & .Fields("cp07"), "" & .Fields("TCT02"), "" & .Fields("TCT03"), "" & .Fields("TCT04"), "" & .Fields("TCT10"), strCC)
                   'end 2018/03/07
                   If strTo <> "" Then
                       'Added by Lydia 2018/03/07
                        If strTo <> "" Then
                            strTo = Mid(strTo, 1, Len(strTo) - 1)
                        End If
                        If strCC <> "" Then
                            strCC = Mid(strCC, 1, Len(strCC) - 1)
                        End If
                       'end 2018/03/07
                       'Added by Lydia 2018/10/31 若停在分案主任一併通知 (ex.FCP-59822)
                       If "" & .Fields("TCT07") <> "" And "" & .Fields("TCT10") = "" Then
                             strTo = strTo & ";" & .Fields("TCT07")
                       End If
                       'end 2018/10/31
                       
                       'Modifed by Lydia 2018/03/07 + strCC
                       PUB_SendMail "QPGMR", strTo, "", strSub, "同主旨", , , , , , strCC
                       strSql = "update TransCaseTitle set TCT115='1'  Where TCT01='" & .Fields("TCT01") & "' "
                       cnnConnection.Execute strSql, intM
                       'Added by Lydia 2018/03/08 記錄log
                       lstHistory.AddItem Now & "--> " & strSub & " 收件者: " & strTo & IIf(strCC <> "", " 副本:" & strCC, "")
                   End If
                   .MoveNext
             Loop
        End With
    End If
    
    Set rsRd = Nothing
'Added by Lydia 2018/03/08
    lstHistory.AddItem Now & "--> 結束 FCP案件命名通知信"
    Exit Sub
    
ErrFCPmail:
    If Err.Number <> 0 Then
       'Modified by Morgan 2019/7/9
       'lstHistory.AddItem Now & "-->FCP案件命名通知信錯誤: " & Err.Description & "-" & Err.Number
       lstHistory.AddItem Now & "-->" & strAct & "錯誤: " & Err.Description & "-" & Err.Number
       'end 2019/7/9
    End If
'end 2018/03/07
End Sub

'Added by Lydia 2023/05/05 外專新案認領-批次處理
Private Sub tmrTFA_Timer()
Dim strR1 As String, intR As Integer
Dim strQ1 As String, intQ As Integer
Dim rsRd As New ADODB.Recordset
Dim rsQD As New ADODB.Recordset
Dim strEDate As String, strETime As String, strNewType As String
Dim strTo As String, strCC As String, strSpecSub As String
Dim bolConn As Boolean
Dim xlsPrintList
Dim wksPrint
Dim strTitle As Variant, strTitleW As Variant
Dim strAct As String
Dim strCont As String

   'Add by Amy 2019/07/11 7點-19點後 才執行
   If Not (Val(Format(Now, "hhmm")) >= Val(strTimeS) And Val(Format(Now, "hhmm")) <= Val(strTimeE)) Then Exit Sub
   If ChkWorkDay(strSrvDate(1)) = False Then
       Exit Sub  '非工作日不執行
   End If
   
   '暫存X分鐘
   tmrMail.Tag = Val(tmrMail.Tag) + 1
   If Val(tmrMail.Tag) < 5 Then Exit Sub
   tmrMail.Tag = 0
   
   If Format(Now, "yyyymmdd") < 外專新案認領啟用日 Then Exit Sub
    
    strAct = "外專新案認領-批次處理"
    If cnnConnection.State <> adStateOpen Then
         lstHistory.AddItem Now & "--> 中斷連線 " & strAct
         WriteLog True
         Exit Sub
    End If
    lstHistory.AddItem Now & "--> 開始 " & strAct

On Error GoTo ErrHandle

   '認領階段(未提申)
   'Modified by Lydia 2023/06/14 (5/19 Email):昨會後經David建議，新案認領組別(非急件)逾期通知 :排除最高主管核判TCN24
   'strR1 = "select tct01,tct04,tcn21,tcn22,tcn23,cp01,cp02,cp03,cp04,pa10 from trackingcasename,transcasetitle,caseprogress,patent " & _
               "where tcn05=tct01(+) and tct04 is null and tct01 is not null and tcn05=cp09(+) and cp159=0 and cp05>=20230501 " & _
               "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and tcn23 in ('0','1','2','3') " & _
               "and (tcn21<to_char(sysdate,'yyyymmdd') or (tcn21=to_char(sysdate,'yyyymmdd') and tcn22<=substr(lpad(to_char(sysdate,'hh24miss'),6,'0'),1,4))) "
   'Modified by Lydia 2025/01/17 暫不認領TCN16=Y的取消日期記錄在命名記錄TransCaseTitle.TCT121,TCT122
   'strR1 = "select tct01,tct04,tcn21,tcn22,cp66,cp67,tcn23,cp01,cp02,cp03,cp04,pa10,tcn25,tcn13,pa75,nvl(fa05,nvl(fa04,fa06)) pa75n,pa26,nvl(cu05,nvl(cu04,cu06)) pa26n " & _
               "From trackingcasename, transcasetitle, caseprogress, patent, fagent, customer " & _
               "where tcn05=tct01(+) and tct04 is null and tct01 is not null and tcn05=cp09(+) and cp159=0 and cp05>=to_char(sysdate,'yyyymmdd')-10000 and cp10 in (" & GetAddStr(FcpAddTct) & ")" & _
               "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and tcn23 in ('0','1','2','3','4','5') and nvl(tcn24,'N') <> 'Y' and nvl(tcn25,0) < 2 " & _
               "and (tcn21<to_char(sysdate,'yyyymmdd') or (tcn21=to_char(sysdate,'yyyymmdd') and tcn22<=substr(lpad(to_char(sysdate,'hh24miss'),6,'0'),1,4))) " & _
               "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
   strR1 = "select tct01,tct04,tcn21,tcn22,NVL(TCT121,cp66) CP66,NVL(TCT122,cp67) CP67,tcn23,cp01,cp02,cp03,cp04,pa10,tcn25,tcn13,pa75,nvl(fa05,nvl(fa04,fa06)) pa75n,pa26,nvl(cu05,nvl(cu04,cu06)) pa26n " & _
               "From trackingcasename, transcasetitle, caseprogress, patent, fagent, customer " & _
               "where tcn05=tct01(+) and tct04 is null and tct01 is not null and tcn05=cp09(+) and cp159=0 and cp05>=to_char(sysdate,'yyyymmdd')-10000 and cp10 in (" & GetAddStr(FcpAddTct) & ")" & _
               "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and tcn23 in ('0','1','2','3','4','5') and nvl(tcn24,'N') <> 'Y' and nvl(tcn25,0) < 2 " & _
               "and (tcn21<to_char(sysdate,'yyyymmdd') or (tcn21=to_char(sysdate,'yyyymmdd') and tcn22<=substr(lpad(to_char(sysdate,'hh24miss'),6,'0'),1,4))) " & _
               "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) "
   strR1 = strR1 & " order by tcn21, tcn22 "
   intR = 1
   Set rsRd = ClsLawReadRstMsg(intR, strR1)
   If intR = 1 Then
      rsRd.MoveFirst
      Do While Not rsRd.EOF
         'Modified by Lydia 2023/06/14 Email主旨開頭改成模組
         strSpecSub = PUB_GetTCNmTitle(rsRd.Fields("cp01"), rsRd.Fields("cp02"), rsRd.Fields("cp03"), rsRd.Fields("cp04"), "" & rsRd.Fields("pa10"), "" & rsRd.Fields("tcn13"), "SPEC")
         strCont = "代理人：" & IIf("" & rsRd.Fields("PA75") <> "", rsRd.Fields("PA75") & " " & rsRd.Fields("PA75N"), "（空白）") & vbCrLf & _
                       "申請人：" & IIf("" & rsRd.Fields("PA26") <> "", rsRd.Fields("PA26") & " " & rsRd.Fields("PA26N"), "（空白）") & vbCrLf
         strNewType = ""
         Select Case "" & rsRd.Fields("tcn23")
             'Modifed by Lydia 2023/06/14
             'Case "0", "2", "3" '急件認領+職代認領(最後)+協調認領
             Case "0"
                 '逾期檢查: 是否有人認領
                 'Modified by Lydia 2023/06/14
                 'If "" & rsRd.Fields("tcn23") = "2" Or "" & rsRd.Fields("tcn23") = "3" Then
                 '    strQ1 = "select st01,st02,tfa05,st16 from transfeeassign,staff where tfa01='" & rsRd.Fields("tct01") & "' and tfa04=st01(+) " & _
                                  "and tfa05='Y' and tfa09=" & CNULL(IIf("" & rsRd.Fields("tcn23") = "2", "1", "2"))
                     strQ1 = "select st01,st02,tfa05,st16 from transfeeassign,staff where tfa01='" & rsRd.Fields("tct01") & "' and tfa04=st01(+) " & _
                                  "and tfa05='Y' and tfa09=" & CNULL(rsRd.Fields("tcn23"))
                 'end 2023/06/14
                     intQ = 1
                     Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
                     If intQ = 1 Then
                        If rsQD.RecordCount = 1 Then
                           bolConn = True
                           cnnConnection.BeginTrans
                               strSql = "Update TrackingCaseName Set TCN20='" & rsQD.Fields("st16") & "' Where TCN05='" & rsRd.Fields("tct01") & "' "
                               cnnConnection.Execute strSql
                               If PUB_UpdateTCNstate("2", rsRd.Fields("cp01") & rsRd.Fields("cp02") & rsRd.Fields("cp03") & rsRd.Fields("cp04"), False) = True Then
                                   lstHistory.AddItem Now & "--> " & rsRd.Fields("cp01") & rsRd.Fields("cp02") & rsRd.Fields("cp03") & rsRd.Fields("cp04") & " 組別: " & rsQD.Fields("st16")
                               Else
                                   lstHistory.AddItem Now & "--> " & rsRd.Fields("cp01") & rsRd.Fields("cp02") & rsRd.Fields("cp03") & rsRd.Fields("cp04") & " 逾期檢查-失敗"
                                   GoTo ErrHandle
                               End If
                           cnnConnection.CommitTrans
                           bolConn = False
                           PUB_SendMailCache True
                           strNewType = "U"
                        End If
                     End If
                 'End If 'Mark by Lydia 2023/06/14
                 
                 '沒人認領或超過1組=>最高主管進行核判TCN24
                 If strNewType = "" Then
                     'strNewType = "4" 'Mark by Lydia 2023/06/14
                     '核判期限至隔日下班前
                     strEDate = CompWorkDay(2, rsRd.Fields("tcn21"))
                     strETime = "1700"
                     'Added by Lydia 2023/06/14
                     strSql = "Update TrackingCaseName Set TCN24='Y', TCN21=" & strEDate & ", TCN22=" & strETime & " Where TCN05='" & rsRd.Fields("tct01") & "' "
                     cnnConnection.Execute strSql
                     'end 2023/06/14
                     'Email通知
                     If PUB_GetTCNEmail(rsRd.Fields("cp01"), rsRd.Fields("cp02"), rsRd.Fields("cp03"), rsRd.Fields("cp04"), IIf(rsRd.Fields("tcn23") = "0", "0", "1")) = True Then
                         lstHistory.AddItem Now & "--> " & rsRd.Fields("cp01") & rsRd.Fields("cp02") & rsRd.Fields("cp03") & rsRd.Fields("cp04") & " 通知最高主管進行核判"
                     Else
                         lstHistory.AddItem Now & "--> " & rsRd.Fields("cp01") & rsRd.Fields("cp02") & rsRd.Fields("cp03") & rsRd.Fields("cp04") & " 通知最高主管進行核判-失敗"
                     End If
                     
                 End If
             Case "1"  '主管認領2H=>職代認領+1H
                 strNewType = "2"
                 Call PUB_CompWorkTime("" & rsRd.Fields("tcn22"), 60, strETime, "" & rsRd.Fields("tcn21"), strEDate)
                 'Email通知參考PUB_UpdateTCNstate
                 strTo = PUB_GetEngGrpMan(strCC)
                 strExc(1) = Replace(strSpecSub, "SPEC", "") & "，請協助確認組別，謝謝！"
                 strExc(2) = strCont
                 PUB_SendMail "QPGMR", strCC, "" & rsRd.Fields("tct01"), strExc(1), strExc(2)
                 lstHistory.AddItem Now & "--> " & rsRd.Fields("cp01") & "-" & rsRd.Fields("cp02") & IIf(rsRd.Fields("cp03") & rsRd.Fields("cp04") <> "000", "-" & rsRd.Fields("cp03") & "-" & rsRd.Fields("cp04"), "") & " 通知職代認領 收件者: " & strCC
             'Added by Lydia 2023/06/14
             Case "2", "3", "4", "5" '已到職代認領(2)或協調(3)、非英說(4)認領階段=> 逾期通知: 1=逾3小時,2=逾1天
                 strExc(3) = ""
                 If Val("" & rsRd.Fields("tcn25")) = 0 Then '以最後認領期限計算:逾3小時
                     Call PUB_CompWorkTime("" & rsRd.Fields("tcn22"), 180, strETime, "" & rsRd.Fields("tcn21"), strEDate)
                     If strSrvDate(1) & Format(Now, "hhmm") >= strEDate & Left(strETime, 4) Then
                        strExc(3) = "1"
                     End If
                 Else '以建檔日期計算:逾1天=> (6/1認領日期+時間逾1天)
                     strExc(2) = CompWorkDay(2, "" & rsRd.Fields("tcn21"))
                     If strSrvDate(1) & Format(Now, "hhmm") >= strExc(2) & Format(rsRd.Fields("tcn22"), "0000") Then
                        strExc(3) = "2"
                     End If
                 End If
                 If strExc(3) <> "" Then
                     strSql = "Update TrackingCaseName Set TCN25='" & strExc(3) & "' Where TCN05='" & rsRd.Fields("TCT01") & "' "
                     cnnConnection.Execute strSql
                     strTo = PUB_GetEngGrpMan(strCC)
                     '同時CC職代+程序
                     strExc(5) = PUB_GetFCPHandler("" & rsRd.Fields("CP01"), "" & rsRd.Fields("CP02"), "" & rsRd.Fields("CP03"), "" & rsRd.Fields("CP04"))
                     strExc(1) = Replace(strSpecSub, "SPEC", IIf(strExc(3) = "1", "-已逾3小時通知", "-已逾24小時通知")) & "，請協助確認組別，謝謝！"
                     strExc(2) = strCont
                     PUB_SendMail "QPGMR", strTo, "" & rsRd.Fields("tct01"), strExc(1), strExc(2), , , , , , strCC & ";" & strExc(5)
                     lstHistory.AddItem Now & "--> " & rsRd.Fields("cp01") & "-" & rsRd.Fields("cp02") & IIf(rsRd.Fields("cp03") & rsRd.Fields("cp04") <> "000", "-" & rsRd.Fields("cp03") & "-" & rsRd.Fields("cp04"), "") & IIf(strExc(3) = "1", "-已逾3小時通知", "-已逾24小時通知") & "  收件者: " & strTo & "  副本: " & strCC & ";" & strExc(5)
                 End If
             'end by Lydia 2023/06/14
         End Select
         If strNewType <> "" And strEDate <> "" And strETime <> "" Then
             strSql = "Update TrackingCaseName Set TCN23='" & strNewType & "' , TCN21=" & CNULL(strEDate, True) & ", TCN22=" & CNULL(Left(strETime, 4), True) & " Where TCN05='" & rsRd.Fields("TCT01") & "' "
             cnnConnection.Execute strSql
         End If
         rsRd.MoveNext
      Loop
   End If
   
   '每個工作日下午２點(14:00)若有前日未核判之新案(非提申急件)，由系統寄email提醒通知國外部最高主管進行核判。
   'Modified by Lydia 2023/05/17 因為放在最後面，1400無法整點執行，改用記錄判斷
   'If Val(Format(Now, "hhmm")) = 1400 Then
   If Val(Format(Now, "hhmm")) >= 1400 And Val(Format(Now, "hhmm")) <= 1430 Then
       strR1 = "select * from addressa4list where aal01='TFA' and aal02=to_char(sysdate,'yyyymmdd') "
       intR = 1
       Set rsRd = ClsLawReadRstMsg(intR, strR1)
       If intR = 0 Then
         strSql = "Delete from addressa4list where aal01='TFA' and aal02<to_char(sysdate,'yyyymmdd') "
         cnnConnection.Execute strSql
         strSql = "Insert into addressa4list (aal01,aal02,aal03,aal04) values ('TFA',to_char(sysdate,'yyyymmdd'),'1','QPGMR') "
         cnnConnection.Execute strSql
   'end 2023/05/17
         strSpecSub = ""
         strTitle = Split("承辦人員,收文日,本所案號,代　理　人,申　請　人,本所期限,法定期限", ",")
         strTitleW = Split("13,10,12,30,30,10,10", ",")
         'Modified by Lydia 2023/06/14 最高主管只管急件逾期或協調不過; and (tcn23='4' or pa10 is null) => and nvl(tcn24,'N')='Y' ,增加 and nvl(tcn23,'0') <> '9'
         'Modified by Lydia 2025/01/17 暫不認領TCN16=Y的取消日期記錄在命名記錄TransCaseTitle.TCT121,TCT122; cp66>>nvl(tct121,cp66)
         strR1 = "select tcn20,tcn21,tcn22,tcn23, tct01,tct02,tct03,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp10,cp13,st02 as cp13n,st16 " & _
                    ",pa75,nvl(fa04,nvl(fa05,fa06)) pa75n,pa26,nvl(cu04,nvl(cu05,cu06)) pa26n " & _
                    "From transcasetitle, caseprogress, staff, trackingcasename, patent, fagent, customer " & _
                    "where tct04 is null and cp159=0 and tct01=cp09(+) and cp13=st01(+) and nvl(tct121,cp66) < to_char(sysdate,'yyyymmdd') and cp05>=to_char(sysdate,'yyyymmdd')-10000 and cp05<" & strSrvDate(1) & _
                    " and tct01=tcn05(+) and nvl(tcn16,'N')<>'Y' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                    " and nvl(tcn24,'N')='Y' and nvl(tcn23,'0') <> '9' and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                    " order by cp05"
         intR = 1
         Set rsRd = ClsLawReadRstMsg(intR, strR1)
         If intR = 1 Then
            rsRd.MoveFirst
            Do While Not rsRd.EOF
               'If "" & rsRd.Fields("tcn23") = "4" Or Not (("" & rsRd.Fields("cp06") <> "" And strSrvDate(1) >= "" & rsRd.Fields("cp06")) Or ("" & rsRd.Fields("cp07") <> "" And strSrvDate(1) >= "" & rsRd.Fields("cp07"))) Then 'Mark by Lydia 2023/06/14
                   If strSpecSub = "" Then
                       strSpecSub = App.path & "\外專新案尚未認領清單" & MsgText(43)
                       If Dir(strSpecSub) <> "" Then
                           Kill strSpecSub
                       End If
                       intQ = 1
                       Set xlsPrintList = CreateObject("Excel.Application")
                       xlsPrintList.SheetsInNewWorkbook = 1
                       xlsPrintList.Workbooks.add
                       Set wksPrint = xlsPrintList.Worksheets(1)
                       wksPrint.Activate
                       'xlsPrintList.Visible = True
                       For intQ = 0 To UBound(strTitle)
                          wksPrint.Range(Chr(65 + intQ) & "1").Value = Trim(strTitle(intQ))
                          wksPrint.Range(Chr(65 + intQ) & ":" & Chr(65 + intQ)).ColumnWidth = Val(strTitleW(intQ))
                       Next intQ
                       wksPrint.Range("A1:" & Chr(65 + UBound(strTitle)) & "1").Font.Bold = True
                       intQ = 2
                   End If
                   wksPrint.Range("A" & intQ).Value = rsRd.Fields("CP13") & " " & rsRd.Fields("cp13n")
                   wksPrint.Range("B" & intQ).Value = ChangeWStringToTDateString(rsRd.Fields("CP05"))
                   wksPrint.Range("C" & intQ).Value = rsRd.Fields("CP01") & "-" & rsRd.Fields("CP02") & IIf(rsRd.Fields("CP03") & rsRd.Fields("CP04") <> "000", "-" & rsRd.Fields("CP03") & "-" & rsRd.Fields("CP04"), "")
                   wksPrint.Range("D" & intQ).Value = rsRd.Fields("PA75") & " " & rsRd.Fields("PA75n")
                   wksPrint.Range("E" & intQ).Value = rsRd.Fields("PA26") & " " & rsRd.Fields("PA26n")
                   wksPrint.Range("F" & intQ).Value = ChangeWStringToTDateString(rsRd.Fields("CP06"))
                   wksPrint.Range("G" & intQ).Value = ChangeWStringToTDateString(rsRd.Fields("CP07"))
               'End If 'Mark by Lydia 2023/06/14
               intQ = intQ + 1
               rsRd.MoveNext
            Loop
            If strSpecSub <> "" Then
               If Val(xlsPrintList.Version) < 12 Then
                   xlsPrintList.Workbooks(1).SaveAs FileName:=strSpecSub, FileFormat:=-4143
               Else
                   xlsPrintList.Workbooks(1).SaveAs FileName:=strSpecSub, FileFormat:=56
               End If
               xlsPrintList.Workbooks.Close
               xlsPrintList.Quit
               Set xlsPrintList = Nothing
               Set wksPrint = Nothing
               strQ1 = Pub_GetSpecMan("外專新案命名核判主管")
               If strQ1 <> "" Then
                   PUB_SendMail "QPGMR", strQ1, "", ChangeTStringToTDateString(strSrvDate(2)) & "外專新案尚未認領清單", "請參考附件", , strSpecSub, , , , , , , , , , , , False
                   lstHistory.AddItem Now & "--> 外專新案尚未認領清單 收件者: " & strQ1
               End If
            End If
         End If
       End If 'Added by Lydia 2023/05/17 改用記錄判斷
   End If
   
   Set rsRd = Nothing
   Set rsQD = Nothing
   lstHistory.AddItem Now & "--> 結束 " & strAct
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      If bolConn = True Then
          cnnConnection.RollbackTrans
      End If
      lstHistory.AddItem Now & "-->" & strAct & "錯誤: " & Err.Description & "-" & Err.Number
   End If

End Sub

'Add by Amy 2019/04/19 電子發票上傳
'Modify by Amy 2024/09/30 常發現已上傳之Tag已更新,又上傳(可能與盟立回傳資料有時間差),故優化
'Modify by Amy 2024/10/23 Acc430上Tag 後將資料搬至History,以利下次檢查回傳檔案,並調整檢查順序
Private Sub tmrTranInv_Timer()
   Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer, iLoop As Integer, intR As Integer, ii As Integer
   Dim bolEndTime As Boolean, bolWSCript As Boolean, bolCheck As Boolean, bolMailToM51 As Boolean, bolToAcc As Boolean, bolData(1) As Boolean '超過時間/WSCRIPT.EXE 已啟用/盟立已回傳檔案/寄信給電腦中心/寄信給上傳者/資料夾有檔案
    Dim strOPath As String, strPath As String, strFileN As String, strMoveErr As String, stUpdErr As String, strSuccess As String, strFail As String '原始資料夾/子資料夾/檔案名/移檔錯誤/更新錯誤/成功檔案/失敗檔案
   Dim strBR03 As String, strTo As String, strToAcc As String, strContent As String, strContent_Fix As String 'BR03人員/錯誤寄信人員/上傳人員/內文/固定內文
   Dim strUpTime As String, strRunMsg As String, strSleep(1) As String, stUpd As String, stPingUrl As String '使用者上傳時間/記錄每個步驟 for 錯誤訊息/暫停時間/更新語法/盟立上傳網址
   Dim intCntS As Integer, intCntF As Integer, intTagCntS As Integer, intTagCntF As Integer  '盟立回傳檔案個數/完成更新個數
   Dim intBookRec As Integer, intTmpCnt As Integer, intLoop As Integer, intReTry1 As Integer, intReTry2 As Integer 'Loop 次數/重試次數/
   Dim strCmd_Fix As String, strDel_Fix As String, strEndTxt_Fix As String, strErr As String
   Dim strTmp As String, strTmp2 As String
    
On Error GoTo ErrHand
   
   '2019/07/22開始 改7點-19點 才執行
   If Not (Val(Format(Now, "hhmm")) >= Val(strTimeS) And Val(Format(Now, "hhmm")) <= Val(strTimeE)) Then Exit Sub

'*** 執行前檢查 ***
   '判斷有連線才做
   If cnnConnection.State <> adStateOpen Then
      strRunMsg = "--> ！！中斷連線 電子發票上傳！！"
      lstHistory.AddItem Now & strRunMsg
      WriteLog True
      Exit Sub
   End If
   
   If ChkBox1.Value = 0 Then Exit Sub
   
   '避免一直發信,只發一次
   strQ = "Select * From BookRecord Where br01=111111 And br02 is not null "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      Exit Sub
   End If
   
'*** 變數設定 ***
   strSleep(0) = 120000: strSleep(1) = 60000
   strTo = "A2004" 'A2004請假,才發信給人事職代 原:"83002;92012;A2004"
   'Memo stToPath 於Form_Load設定
   strOPath = "C:\551cron\"
  
   '2020/01/21 增加 Ping 對方網址是否是通的-薛經理
   stPingUrl = "20.239.56.83"
   intLoop = 5
   '2019/12/29 加 測式資料庫 連盟立[測式]平台
   If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 And ChkBox1.Value = 1 Then
      strSleep(0) = 6000: strSleep(1) = 3000
      stPingUrl = "xmltest.551.com.tw"
      If strUserNum <> "A2004" Then strRunExePC = "\\A2004\"
      intLoop = 2
   Else
      strRunExePC = "\\" & pub_HostName & "\"
   End If
   
   strDel_Fix = "Delete BookRecord Where br01=111111 "
   strCmd_Fix = "Update BookRecord Set br02=" & strSrvDate(1) & " Where br01=111111 "
   strContent_Fix = "Update BookRecord Set br02=null Where br01=111111 " & vbCrLf & _
                                    "！！電子發票上傳才能再執行！！"
   strEndTxt_Fix = "--> 結束　電子發票上傳"

'*** End 變數設定 ***
  
   'C:\551cron\Success\當天日..\ 仍有檔案
   strTmp = strOPath & "Success\"
   bolData(0) = ChkXml(0, strTmp)
   'C:\551cron\Error\當天日..\ 仍有檔案
   strTmp2 = strOPath & "Error\"
   bolData(1) = ChkXml(0, strTmp2)
   
   '讀取 stToPath=C:\Einvoice\ 資料夾檔案
   strRunMsg = "--> 1.讀取資料夾[" & stToPath & "]"
   File1.path = stToPath
   File1.Refresh
   '因 2020/01/06 Success資料夾已有上傳成功資料,但UpCast資料有殘留資料,故加判斷
   File2.path = strOPath & "UpCast" 'C:\551cron\UpCast
   File2.Refresh
   
   '判斷財務有上傳Tag (於Account-11t0新增其Tag)
   intBookRec = 0
   strQ = "Select * From BookRecord Where br01=111111 "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      intBookRec = RsQ.RecordCount
      strUpTime = RsQ.Fields("BR05")
      strBR03 = "" & RsQ.Fields("BR03") 'BookRecord無Tag,仍有未上傳資料 用
   End If
   
   '[無] BookRecord Tag,發mail通知電腦中心
   If intBookRec = 0 And (bolData(0) = True Or bolData(1) = True Or File1.ListCount > 0 Or File2.ListCount > 0) Then
      If bolData(0) = True Then 'Success 資料夾
         strErr = Replace(strTmp, "C:", "C$") & " 有資料"
      ElseIf bolData(1) = True Then 'Error 資料夾
         strErr = Replace(strTmp2, "C:", "C$") & " 有資料"
      ElseIf File2.ListCount > 0 Then 'UpCast 資料夾
         strErr = Replace(strOPath, "C:", "C$") & "UpCast 資料夾中有 [" & File2.ListCount & "] 個檔案"
      Else
         'C:\Einvoice\ 資料夾
         strErr = Replace(strOPath, "C:", "C$") & " 資料夾中有 [" & File1.ListCount & "] 個檔案"
      End If
      strErr = pub_HostName & "\" & strErr & "未上傳"
      strTmp = ServerTime
      stUpd = "Insert into BookRecord(br01,br03,br02,br04,br05) Values(111111,'" & strTo & "'," & strSrvDate(1) & "," & strSrvDate(1) & "," & strTmp & ")"
      cnnConnection.Execute stUpd
      
      strRunMsg = "BookRecord無Tag,但有資料未上傳"
      strContent = Replace(strRunMsg, "但有資料未上傳", "但 " & strErr & "未上傳") & vbCrLf & _
                              "請確認原因，排除問題後請執行 " & vbCrLf & strContent_Fix
      PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strContent, , , , , , , "QPGMR", , , , False
      Exit Sub
   '[有] BookRecord Tag 且 Success 或 Error 資料夾有資料
   ElseIf intBookRec = 1 And (bolData(0) = True Or bolData(1) = True) Then
      bolCheck = True
      GoTo UpdTag
   '[有] BookRecord Tag 且 無要上傳資料
   ElseIf intBookRec = 1 And File1.ListCount = 0 And File2.ListCount = 0 Then
      cnnConnection.Execute strCmd_Fix
      
      strRunMsg = "BookRecord有Tag,但無資料需上傳"
      strContent = strRunMsg & "，請確認原因" & vbCrLf & _
                              "排除問題後" & vbCrLf & _
                              "1.[有]資料要上傳 ,請執行[ " & Replace(strContent_Fix, vbCrLf & "！！電子發票上傳才能再執行！！", "") & "]" & vbCrLf & _
                              "2.[無]資料要上傳 ,請執行[ " & strDel_Fix & "]" & vbCrLf & _
                              "！！電子發票上傳才能正常執行！！"
      PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strContent, , , , , , , "QPGMR", , , , False
      Exit Sub
   ElseIf intBookRec = 0 And File1.ListCount = 0 And File2.ListCount = 0 Then
      Exit Sub
   End If

   If ChkPing(stPingUrl) = -1 Then
      cnnConnection.Execute strCmd_Fix
            
      strRunMsg = "！！Ping 盟立網址不通，請與盟立聯絡！！"
      strContent = "如主旨，確認盟立網址已通，請執行下列語法：" & vbCrLf & strContent_Fix
      
      PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strContent, , , , , , , "QPGMR", , , , False
      lstHistory.AddItem Now & "--> " & strRunMsg
      WriteLog True
      Exit Sub
   End If
'*** End 執行前檢查 ***
   
   strRunMsg = "--> 開始　電子發票上傳"
   lstHistory.AddItem Now & strRunMsg
   
   strContent = "": strErr = ""
   '已有 WSCRIPT.EXE在執行
   
   If PUB_CheckIsRunning("WSCRIPT.EXE") = True Then
      cnnConnection.Execute strCmd_Fix
      
      strRunMsg = "！！有WSCRIPT.EXE在執行，確認是否當掉！！"
      strContent = "如主旨，排除問題後，請執行下列語法：" & vbCrLf & strContent_Fix
            
      PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strContent, , , , , , , "QPGMR", , , , False
      lstHistory.AddItem Now & "--> " & strRunMsg
      lstHistory.AddItem Now & strEndTxt_Fix
      WriteLog True
      Exit Sub
   '目前沒在 Run WSCRIPT.EXE,啟動WSCRIPT.EXE
   Else
      bolWSCript = True
      ChDir strOPath 'strOPath="C:\551cron\"
      Shell "WSCRIPT.EXE move_start.vbs"
      Shell "WSCRIPT.EXE start.vbs"
      ChDir App.path
      strRunMsg = "--> 1.開始執行Wscript.exe-->OK"
      DoEvents
      Sleep Val(strSleep(0)) '先等待WSCRIPT.EXE關閉
   End If

   '已執行 WSCRIPT.EXE
   If bolWSCript = True Then
TryAgin:
      'WSCRIPT.EXE[已]關閉
      If PUB_CheckIsRunning("WSCRIPT.EXE") = False Then
         iLoop = 0
TryAgin2:
         Do While iLoop < intLoop
            iLoop = iLoop + 1
            '判斷使用者上傳資料夾內是否有資料
            If Dir(stToPath & "*.*") = MsgText(601) Then
               '判斷上傳中繼資料夾內是否有資料
               If Dir(strOPath & "UpCast\*.*") = MsgText(601) Then
                  '加檢查 Success,避免檔案已上傳至盟立,但盟立未回傳檔案,而更新Tag
                  If ChkXml(0, strOPath & "Success\") = True Then
                     bolCheck = True
                     Exit Do
                  End If
                  '加檢查 Error,避免檔案已上傳至盟立,但盟立未回傳檔案,而更新Tag
                  If ChkXml(0, strOPath & "Error\") = True Then
                     bolCheck = True
                     Exit Do
                  End If
               End If
            End If
            DoEvents
            Sleep Val(strSleep(1)) '先等待再判斷是否上傳成功
         Loop
         
         '盟立仍未回傳資料
         If bolCheck = False Then
            '已重試但User 上傳後已超過20分鐘
            bolEndTime = ChkRunTime(strUpTime, "20")
            If (intReTry2 > 0 And bolEndTime = True) Or intReTry2 >= 1 Then
               bolMailToM51 = True
               
               strRunMsg = "盟立仍未回傳資料"
               strErr = "，已重試[" & intReTry2 & "]次"
               If bolEndTime = True Then
                  strErr = strErr & "，但User 上傳後已超過20分鐘"
               End If
               strErr = strRunMsg & strErr
               
               strContent = strErr & vbCrLf & _
                                    "請確協助確認原因" & vbCrLf & _
                                    "問題排除後,仍[有]資料要上傳 ,請執行下列語法： " & vbCrLf & strContent_Fix
            '未重試
            ElseIf intReTry2 = 0 Then
               intReTry2 = intReTry2 + 1
               strRunMsg = "        -->盟立仍未回傳資料,重試：[" & intReTry2 & "]"
               lstHistory.AddItem Now & strRunMsg
               GoTo TryAgin2
            End If
         End If
      'WSCRIPT.EXE[未]關閉
      Else
         '已重試,已Run 超過20分鐘
         bolEndTime = ChkRunTime(strUpTime, "20")
         If (intReTry1 > 0 And bolEndTime = True) Or intReTry1 >= 2 Then
            bolMailToM51 = True
            
            strRunMsg = "WSCRIPT.EXE[未]關閉"
            strErr = "，已重試[" & intReTry2 & "]次"
            If bolEndTime = True Then
               strErr = strErr & "，但User 上傳後已超過20分鐘"
            End If
            strErr = strRunMsg & strErr
            
            strContent = strErr & vbCrLf & _
                                    "請確協助確認 WSCRIPT.EXE 是否當掉！！" & vbCrLf & _
                                    "問題排除後,仍[有]資料要上傳 ,請執行下列語法： " & vbCrLf & strContent_Fix
         '未重試
         Else
            If intReTry1 = 0 Then
               intReTry1 = intReTry1 + 1
               DoEvents
               Sleep Val(strSleep(0)) * 2 '先等待WSCRIPT.EXE關閉
            Else
               intReTry1 = intReTry1 + 1
            End If
            strRunMsg = "        -->WSCRIPT.EXE仍在執行中,重試：[" & intReTry1 & "]"
            lstHistory.AddItem Now & strRunMsg
            GoTo TryAgin
         End If
      End If
      If bolMailToM51 = True Then
         cnnConnection.Execute strCmd_Fix
            
         PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strContent, , , , , , , "QPGMR", , , , False
         lstHistory.AddItem Now & "--> 電子發票上傳有誤！已發信通知相關人員"
         lstHistory.AddItem Now & strEndTxt_Fix
         Exit Sub
      End If
      
UpdTag:
      strContent = "": strErr = ""
      bolMailToM51 = False
      strPath = strOPath & "Success\" & Mid(strSrvDate(1), 1, 4) & "\" & Mid(strSrvDate(1), 5, 2) & "\" & Mid(strSrvDate(1), 7, 2) & "\"
      If bolCheck = True Then
      '*** 最後更新Tag及發信 ***
         intCntS = 0: intCntF = 0: intTagCntS = 0: intTagCntF = 0
         strSuccess = "": strFail = ""
         '讀取此次上傳資料,以發票號排
         strQ = "Select AccTmp11t0.*,a4319,a4320,a4321,a4322,a4324,a4325 From AccTmp11t0,Acc430 " & _
                     "Where R004=" & strUpTime & " And R001 is null And SubStr(R002,5,10)=a4301(+) " & _
                     "Order by SubStr(R002,5,10) "
         intQ = 1
         Set RsQ = ClsLawReadRstMsg(intQ, strQ)
         If intQ = 1 Then
            strRunMsg = "--> 2.更新Tag及移檔-->"
            RsQ.MoveFirst '1090120 同一筆辜上傳二次,一筆成功一筆失敗
            intTmpCnt = RsQ.RecordCount
            strToAcc = RsQ.Fields("ID") '上傳人員
            Do While RsQ.EOF = False
               strMoveErr = "": stUpdErr = ""
               strFileN = RsQ.Fields("R002")
               'Success 資料夾有上傳之檔案
               If Dir(strPath & strFileN & "*.*") <> MsgText(601) Then
                  strSuccess = strSuccess & ";" & strFileN
                  intCntS = intCntS + 1
                  If MoveFile(False, strOPath & "Success\", strFileN & "*.*", strMoveErr) = False Then
                     bolMailToM51 = True
                     strErr = strErr & ";" & strMoveErr
                  '移檔成功才更新Tag
                  Else
                     If UpdInvTag(0, strFileN, stUpdErr) = False Then
                        bolMailToM51 = True
                        strErr = strErr & ";" & stUpdErr
                        '更新有誤,檔案搬回
                        If MoveFile(True, strOPath & "Success\History\", strFileN & "*.*", strMoveErr) = False Then
                           strErr = strErr & "且" & strMoveErr & "(搬回)"
                        End If
                     Else
                        intTagCntS = intTagCntS + 1
                     End If
                  End If
               'Error 資料夾有上傳之檔案
               ElseIf Dir(Replace(strPath, "\Success\", "\Error\") & strFileN & "*.*") <> MsgText(601) Then
                  strFail = strFail & ";" & strFileN
                  intCntF = intCntF + 1
                  If MoveFile(False, strOPath & "Error\", strFileN & "*.*", strMoveErr) = False Then
                     bolMailToM51 = True
                     strErr = strErr & ";" & strMoveErr
                  Else
                     If UpdInvTag(1, strFileN, stUpdErr) = False Then
                        bolMailToM51 = True
                        strErr = strErr & ";" & stUpdErr
                        '更新有誤,檔案搬回
                        If MoveFile(True, strOPath & "Error\History\", strFileN & "*.*", strMoveErr) = False Then
                           strErr = strErr & "且" & strMoveErr & "(搬回)"
                        End If
                     Else
                        intTagCntF = intTagCntF + 1
                     End If
                  End If
               End If
               RsQ.MoveNext
            Loop
            '無任何[成功]與[失敗]資料
            If intCntS = 0 And intCntF = 0 Then
               strRunMsg = strRunMsg & "盟立未回傳任何資料"
               cnnConnection.Execute strCmd_Fix
               
               strRunMsg = "--> 上傳共[" & intTmpCnt & "]筆，但無任何[成功]與[失敗]筆數"
               strContent = "請協助確認是何問題！！" & vbCrLf & _
                                       "排除問題後,仍[有]資料要上傳,請執行下列語法： " & vbCrLf & strContent_Fix
               
               PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strContent, , , , , , , "QPGMR", , , , False
               lstHistory.AddItem Now & strRunMsg & "->已發信通知電腦中心"
               lstHistory.AddItem Now & strEndTxt_Fix
               Exit Sub
            End If
           
            strContent = "此次上傳共[" & intTmpCnt & "]筆" & vbCrLf & _
                                       "成功：[" & intCntS & "]筆" & vbCrLf & _
                                       "失敗：[" & intCntF & "]筆" & vbCrLf
                                       
            '有上傳成功資料 Or 全數上傳失敗,發信通知上傳者
            If intTmpCnt > 0 And (intCntS > 0 Or intTmpCnt = intCntF) Then
               If strToAcc = MsgText(601) Then strToAcc = strTo
               PUB_SendMail "QPGMR", strToAcc, "", "電子發票上傳已完成，請至盟立平台確認！", strContent, , , , , , , "QPGMR", , , , False
               bolToAcc = True
               '全數上傳失敗,發信通知電腦中心
               If intTmpCnt = intCntF Then bolMailToM51 = True
            End If
            '全數成功 or 全數失敗(包含移檔及Tag更新)
            If intTmpCnt > 0 And (intTmpCnt = intTagCntS Or intTmpCnt = intTagCntF) Then
               strRunMsg = strRunMsg & "OK"
               lstHistory.AddItem Now & strRunMsg
            ElseIf intTmpCnt > 0 And strErr <> MsgText(601) Then
               strRunMsg = strRunMsg & "有誤"
               lstHistory.AddItem Now & strRunMsg
            End If
            '[暫存檔]Tag更新
            strRunMsg = "--> 3.刪除BookRecord-->"
            If Pub_GetField("AccTmp11t0", "R004=" & strUpTime & " And R001 is null", "Distinct R004") = MsgText(601) Then
               '刪除BookRecord
               cnnConnection.Execute strDel_Fix, intR
               If intR = 0 Then
                  bolMailToM51 = True
                  strRunMsg = strRunMsg & "！！失敗！！"
                  strErr = strErr & ";" & strQ & "-->此語法[無]資料可更新" & vbCrLf
               Else
                  strRunMsg = strRunMsg & "OK"
               End If
               lstHistory.AddItem Now & strRunMsg
            '[暫存檔]Tag有未更新
            Else
               strRunMsg = strRunMsg & " 暫存檔 有Tag未更新"
               cnnConnection.Execute strCmd_Fix
               
               bolMailToM51 = True
               strErr = strErr & ";" & strRunMsg & "-->不會刪BookRecord" & vbCrLf
               lstHistory.AddItem Now & strRunMsg
            End If
            '發信電腦中心
            If strErr <> MsgText(601) Or bolMailToM51 = True Then
               If bolToAcc = True Then
                  strContent = strContent & "<以上為已通知財務之信件內文>" & vbCrLf & vbCrLf
               End If
               strContent = strContent & "此次Tag更新如下" & vbCrLf & _
                                       "成功：共 [" & intTagCntS & "]筆" & vbCrLf & _
                                       IIf(strSuccess = "", "", Replace(Mid(strSuccess, 2), ";", vbCrLf) & vbCrLf) & vbCrLf & _
                                       "失敗：共 [" & intTagCntF & "]筆" & vbCrLf & _
                                       IIf(strFail = "", "", Replace(Mid(strFail, 2), ";", vbCrLf) & vbCrLf) & vbCrLf & vbCrLf
               strContent = strContent & "<電腦中心需協助確認及處理如下>" & Replace(strErr, ";", vbCrLf) & vbCrLf
               '非全數失敗,加顯示文字
               If Not (intTmpCnt > 0 And intTmpCnt = intCntF) Then
                  strContent = strContent & "排除問題後,仍[有]資料要上傳,請執行下列語法： " & vbCrLf & strContent_Fix
               End If
               
               PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strContent, , , , , , , "QPGMR", , , , False
               lstHistory.AddItem Now & "--> 電子發票上傳有誤！已發信通知電腦中心"
            End If
         End If 'intQ = 1
      '*** End 最後更新Tag及發信 ***
      End If
   End If 'bolWSCript = True
   
   strRunMsg = strEndTxt_Fix
   lstHistory.AddItem Now & strRunMsg
   Exit Sub
    
ErrHand:
   If strRunMsg <> MsgText(601) Then
      strExc(9) = "": strErr = ""
      
      If Err.Number = 76 Then
         strExc(9) = "請確認資料夾[權限][共用]頁籤權限已開"
      Else
          strExc(9) = Err.Description
      End If
      
      cnnConnection.Execute strCmd_Fix
      strErr = "排除問題後,仍[有]資料要上傳,請執行下列語法： " & vbCrLf & strContent_Fix
      strErr = strRunMsg & "有誤！！-->" & strExc(9) & vbCrLf & strErr
                     
      PUB_SendMail "QPGMR", strTo, "", "電子發票上傳有誤,請確認！(" & strRunMsg & ")", strErr, , , , , , , "QPGMR", , , , False
   End If
   lstHistory.AddItem Now & "--> 電子發票上傳有誤！" & strExc(9) & vbCrLf & "步驟：" & strRunMsg
End Sub

'更新發票檔及折讓上傳Tag
'Modify by Amy 2019/07/22
'Modify by Amy 2024/10/23 改為單筆更新並加 AccTmp11t0
'intChoose:0-更新成功Tag/更新失敗Tag
Private Function UpdInvTag(intChoose As Integer, ByVal stInvNo As String, ByRef stErr As String) As Boolean
   Dim stUpd As String, stUpd2 As String, stTBN As String, stField1 As String, stField2 As String, stKey As String, stTmp As String
   Dim bolTrans As Boolean, j As Integer, intRun As Integer, intRun2 As Integer
   Dim arrTmp
 On Error GoTo ErrHand
 
   UpdInvTag = False
   stErr = ""
   stTmp = stInvNo
   
   '成功
   If intChoose = 0 Then
      '依發票號碼及上傳代碼更新上傳[Acc430]及更新暫存檔[AccTmp11t0]Tag
      stUpd = "":  stTBN = "": stField1 = "": stField2 = ""
      stKey = Mid(stTmp, Val(InStr(stTmp, "_")) + 1)
      Select Case Left(stTmp, 3)
         Case "A04", "C04" '開立發票
            stTBN = "Acc430": stUpd = "a4301='" & stKey & "' And a4319 is null "
            stField1 = "a4319": stField2 = "a4320"
         Case "A05", "C05" '作廢發票
            stTBN = "Acc430": stUpd = "a4301='" & stKey & "' And a4321 is null "
            stField1 = "a4321": stField2 = "a4322"
         Case "B04", "D04" '折讓開立
            'Modify by Amy 2019/12/27 +if 第5碼是I編號(真正折讓),條件拿掉 And a0s04='3'/第5碼不是I編號(非真正折讓)先給客戶發票但未付款,已申報後轉開
            If Mid(stTmp, 5, 1) = "I" Then
               stTBN = "Acc0S0": stUpd = "a0s01='" & stKey & "' And a0s26 is not null And a0s28 is null "
               stField1 = "a0s28": stField2 = "a0s29"
            Else
               stTBN = "Acc430": stUpd = "a4301='" & stKey & "' And a4310 is not null And a4324 is null "
               'Modify by Amy 2020/07/16 bug 欄位錯誤
               stField1 = "a4324": stField2 = "a4325"
            End If
            '折讓作廢目前沒有
      End Select
      
      stUpd = "Update " & stTBN & _
                     " Set " & stField1 & "=" & Val(strSrvDate(2)) & "," & stField2 & "=" & Format(Now, "HHmmss") & _
                     " Where " & stUpd
      stUpd2 = "Update AccTmp11t0 Set R001='S' Where R002='" & stTmp & "' "
      
      cnnConnection.BeginTrans
      bolTrans = True
      cnnConnection.Execute stUpd, intRun
      cnnConnection.Execute stUpd2, intRun2
      cnnConnection.CommitTrans
      
      If intRun = 0 Then
         stErr = stUpd & "-->此語法[無]資料可更新！(發票號:" & stKey & ")"
      End If
      If intRun2 = 0 Then
         stErr = stUpd2 & "-->此語法[無]資料可更新！(發票檔案:" & stTmp & ")"
      End If
   '失敗
   Else
      stUpd = "Update AccTmp11t0 Set R001='F' Where R002='" & stTmp & "' "
      cnnConnection.Execute stUpd, intRun
      If intRun = 0 Then
         stErr = stUpd & "-->此語法[無]資料可更新！(發票檔案:" & stTmp & ")"
      End If
   End If
   
   UpdInvTag = True
   Exit Function
    
ErrHand:
   If Err.Number <> 0 Then
      If bolTrans = True Then
         cnnConnection.RollbackTrans
         stErr = "更新資料表有誤(" & Err.Description & ")！" & vbCrLf & _
                     "語法：" & stUpd
      Else
          stErr = Err.Description
      End If
      stErr = stErr & vbCrLf & "發票檔案：" & stTmp
   End If
End Function
'end 2019/7/11

Public Function ChkPing(stMachine) As Long
    Dim ObjPing, objStatus, no
    
    If pub_OS = 2 Then
        Set ObjPing = Interaction.GetObject("winmgmts:").ExecQuery("select * from Win32_PingStatus where address = '" & stMachine & "'")
        For Each objStatus In ObjPing
            no = IsNull(objStatus.StatusCode) Or objStatus.StatusCode <> 0
            ChkPing = IIf(no, -1, Abs(objStatus.ResponseTime))
        Next
        Set ObjPing = Nothing
    End If
End Function

'Added by Lydia 2020/03/19 外專利益衝突搬檔
Private Sub tmrMoveList_Timer()
Dim strDate As String, strTime  As String, strAct As String
Dim intP As Integer, intK As Integer, intU As Integer
Dim strNewName As String
Dim strDefPath As String
Dim strPass As String
Dim fs, fso, fl
Dim rsAD As New ADODB.Recordset
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim strKey As String
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, m_CP10 As String
Dim nCP09 As String 'D類收文號
Dim m_TempDir As String 'Key之前的路徑
Dim m_TempName As String  '統一名稱
Dim nMax As Integer
Dim nPos As String
Dim strErr As String, strErrCont As String
Dim strCont As String
Dim iCount As Long
Dim intB As Integer, rsB As New ADODB.Recordset 'Added by Lydia 2020/03/24

     Exit Sub 'Added by Lydia 2020/04/10 已搬檔完成，先取消執行
     
'-----Test 2020/03/02
'Memo by Lydia 2020/03/18 設定搬檔作業
'搬檔時間：每天20:00~次日01:59之間 ; DB備份從每日02:00開始
'設定批次:
'1.將要搬檔的第一層資料夾路徑放在Lydia_movelist(目前整理到20200317的資料夾路徑)；本機端測試放在Lydia_a002
'2.設定當天要搬擋的記錄R004=系統日
'3.掛在TeAutoPdf程式，每天20:00開始執行(正式程式放在\\m51-win7\c$\taie\AutoPdf, 複製專案到本機端)
'4.執行完成後會發完成通知email，若有無法處理的檔案會寫在email內文。
'---Lydia_MoveList的欄位
'序號R001 Number(4)
'路徑R002 VARCHAR2(100)
'建立日期R003 VARCHAR2(8)
'預定日期R004 VARCHAR2(8)
'執行日期R005 VARCHAR2(8)
'例外執行日期R006 VARCHAR2(8)
'end 2020/03/18

On Error GoTo ExceptCont

    '暫存X分鐘
    tmrMoveList.Tag = Val(tmrMoveList.Tag) + 1
    If Val(tmrMoveList.Tag) < 5 Then Exit Sub
    tmrMoveList.Tag = 0

    strAct = "外專利益衝突搬檔"
   '用machine Time
    strDate = Format(Now, "yyyymmdd")
    strTime = Format(Now, "hhmm")
    '上班時間控制啟動作業
    strExc(1) = Pub_GetSpecMan("XY搬檔指定日")
    If strExc(1) <> "" And InStr(strExc(1), strDate) > 0 Then
        If bolMoveStatus = True Then Exit Sub 'Added by Lydia 2020/03/25
    Else
        '啟動時間: 每天20:00~次日01:59之間
        'Modified by Lydia 2020/03/23 每天20:00開始
        'If Not (strTime > "2000" Or strTime < "0200") Or bolMoveStatus = True Then
        'Modified by Lydia 2020/03/25 增加跨日
        'If strTime < "2000" Or bolMoveStatus = True Then
        If Not (strTime >= "2000" Or (strTime >= "0100" And strTime <= "0210")) Or bolMoveStatus = True Then
            Exit Sub
        End If
    End If
    
    'Added by Lydia 2020/03/25 跨日抓前一天未完成
    If strTime >= "0100" And strTime <= "0210" Then
          strDate = CompDate(2, -1, strDate)
    End If
    'end 2020/03/25
    
    '判斷有連線才做
    If cnnConnection.State <> adStateOpen Then
         lstHistory.AddItem Now & "--> 中斷連線 外專利益衝突搬檔"
         WriteLog True
         Exit Sub
    End If
    lstHistory.AddItem Now & "--> 開始 外專利益衝突搬檔"
    
    bolMoveStatus = True
    
    '注意跨日作業:
    strExc(0) = "select r002 from Lydia_MoveList where r005 is null and r004= " & strDate & " order by r001 "
    strExc(0) = "select * from (" & strExc(0) & ") where rownum < 11 order by 1" 'Added by Lydia 2020/03/25
    intI = 1
    Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        strCont = "" & rsAD.GetString(adClipString, , , vbCrLf)
    End If
    strCont = "搬檔作業：" & strCont
    
    'Table 只記錄第一層資料夾路徑, 之後直接抓目前檔案
    strExc(0) = "select * from Lydia_MoveList where r005 is null and r004= " & strDate & " order by r001 "
    strExc(0) = "select * from (" & strExc(0) & ") where rownum < 11 order by 1" 'Added by Lydia 2020/03/25 因為24日晚上只上傳410~411就疑似CUP滿載,所以分段執行
    intI = 1
    Set rsAD = ClsLawReadRstMsg(intI, strExc(0))
    If intI = 1 Then
        Set fs = CreateObject("Scripting.FileSystemObject") 'Added by Lydia 2020/03/23
        strCont = strCont & vbCrLf & "執行時間：" & ChangeWStringToWDateString(strSrvDate(1)) & "(" & Format(ServerTime, "00:00:00") & ")"
        With rsAD
            .MoveFirst
            Do While Not .EOF
                m_TempDir = ""
                strDefPath = "" & .Fields("r002")
                nPos = "" & .Fields("r001")
                If Right(strDefPath, 1) <> "\" Then strDefPath = strDefPath & "\" '與抓子資料夾有關
                
'English_vers分析：本所案號、上傳類型
                strKey = "\ENGLISH_VERS\"
                intK = InStr(UCase(strDefPath), strKey)
                If intK > 0 Then
                    m_CP10 = cntEnglish_Vers
                    strExc(1) = Mid(strDefPath, intK + Len(strKey))
                    tmpArr1 = Split(strExc(1), "\") '用\區隔路徑層級
                    If Len(tmpArr1(0)) = 6 Then  'FMP案
                        m_CP01 = "P"
                        m_CP02 = tmpArr1(0)
                    ElseIf Len(tmpArr1(0)) = 3 Then 'FCP案
                        m_CP01 = "FCP"
                        m_CP02 = tmpArr1(0)
                    End If
                    m_CP03 = "0": m_CP04 = "00"
                
                    If m_CP01 = "FCP" And Len(m_CP01 & m_CP02 & m_CP03 & m_CP04) < 12 Then
                        '先將所有案號資料夾記錄在字串
                        strExc(9) = ""
                        strPass = Dir(strDefPath, vbDirectory)
                        Do While strPass <> ""
                             If strPass <> "." And strPass <> ".." Then
                                 If GetAttr(strDefPath & strPass) = vbDirectory Then
                                      If Val(strPass) > 10000 And Val(strPass) < 64000 And InStr(strDefPath, Left(strPass, 3)) > 0 Then 'Added by Lydia 2020/03/23 判斷是否為5碼本所案號
                                          strExc(9) = strExc(9) & "," & strPass
                                      End If 'Added by Lydia 2020/03/23
                                 End If
                             End If
                             strPass = Dir()
                        Loop
                        If strExc(9) <> "" Then
                            tmpArr1 = Empty
                            tmpArr1 = Split(Mid(strExc(9), 2), ",")
                            nMax = UBound(tmpArr1) + 1
                            m_CP02 = Format(tmpArr1(0), "000000")
                            m_TempDir = strDefPath & tmpArr1(0) & "\"
                        End If
                    ElseIf m_CP01 = "P" And Len(m_CP01 & m_CP02 & m_CP03 & m_CP04) = 10 Then
                         nMax = 1
                         m_TempDir = strDefPath
                    End If
                    
                    '逐案號資料夾上傳
                    For intI = 1 To nMax
                        If intI > 1 Then
                            m_CP02 = Format(tmpArr1(intI - 1), "000000")
                            m_TempDir = strDefPath & tmpArr1(intI - 1) & "\"
                        End If
                        nCP09 = ""
                        intP = 0
                        '1.先讀取資料夾的所有檔案，拿掉檔名的Unicode字
                         'Set fs = CreateObject("Scripting.FileSystemObject") 'Remove by Lydia 2020/03/23
                         If Trim(m_TempDir) = "" Then GoTo JumpTo01
                         If Not fs.FolderExists(m_TempDir) Then
                            strErrCont = strErrCont & vbCrLf & "資料夾不存在：" & m_TempDir
                            GoTo JumpTo01
                         Else
                            Set fso = fs.GetFolder(m_TempDir)
                            For Each fl In fso.files
                               TxtFile.Text = fl.Name
                               strNewName = TxtFile.Text 'Added by Lydia 2020/03/27
                               strErr = "拿掉Unicode字:" & TxtFile.Text
                               If TxtFile.Text <> fl.Name Then
                                   strNewName = Replace(TxtFile.Text, "?", "x")
                                   'Added by Lydia 2020/03/26 檢查是否有相同新檔名
                                   If ChkNewFileName(m_TempDir, strNewName) = False Then
                                        strErrCont = strErrCont & vbCrLf & "有相同新檔名:" & m_TempDir & TxtFile.Text & "，請確認檔名！"
                                        GoTo JumpTo01
                                   Else
                                   'end 2020/03/26
                                        fl.Name = strNewName
                                        TxtFile.Text = strNewName 'Added by Lydia 2020/03/27
                                   End If 'end 2020/03/26
                                   
                               'Added by Lydia 2020/03/27 整合更名
                               Else
                                    If ChkNewFileName(m_TempDir, strNewName) = False Then
                                        strErrCont = strErrCont & vbCrLf & "有相同新檔名:" & m_TempDir & TxtFile.Text & "，請確認檔名！"
                                        GoTo JumpTo01
                                    Else
                                         If strNewName <> TxtFile.Text Then
                                              fl.Name = strNewName
                                         End If
                                    End If
                               'end 2020/03/27
                               End If
                            Next
                         End If
                         'Added by Lydia 2020/03/24 先檢查本所案號是否存在
                         If Len(m_CP02) = 6 Then
                              strExc(0) = "select pa01,pa02,pa03,pa04 from patent where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='0' and pa04='00' "
                              intB = 1
                              Set rsB = ClsLawReadRstMsg(intB, strExc(0))
                              If intB = 0 Then
                                strErrCont = strErrCont & vbCrLf & "無基本檔:" & m_TempDir & "，請確認檔名是否正確！"
                                GoTo JumpTo01
                              End If
                         End If
                         
                         '2.上傳檔案，若有子資料夾則壓縮為.zip
JumpToReDir:
                         strPass = Dir(m_TempDir, vbDirectory)
                         intP = 0: intU = 0
                         Do While strPass <> ""
                            If strPass <> "." And strPass <> ".." Then
                                TxtFile.Text = strPass
                                If InStr(TxtFile.Text, ".") = 0 And InStr(TxtFile.Text, "?") > 0 Then '子資料夾有Unicode
                                    strErrCont = strErrCont & vbCrLf & "子資料夾名稱有Unicode: " & m_TempDir & TxtFile.Text
                                    intU = intU + 1
                                    GoTo JumpToPass
                                End If
                                strNewName = ""
                                If GetAttr(m_TempDir & strPass) = vbDirectory Then
                                    strNewName = m_TempDir & Mid(strPass, 1, 20) & "." & Format(FileDateTime(m_TempDir & strPass), "YYYYMMDDHHMMSS")
                                    If ZipFolder(m_TempDir & strPass, strNewName) = True Then
                                          strNewName = strNewName & ".zip"
                                          Call PUB_KillTempFolder(strPass, m_TempDir)
                                          Sleep 100 'Added by Lydia 2020/03/25 休眠1秒
                                          'Added by Lydia 2020/03/25 判斷子資料夾是否存在
                                          If Not fs.FolderExists(m_TempDir & strPass) Then
                                                intP = 0: intU = 0 '因為重讀資料夾,計數歸0
                                                GoTo JumpToReDir  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到
                                          'Added by Lydia 2020/03/25
                                          Else
                                                strErrCont = strErrCont & vbCrLf & "子資料夾未刪: " & m_TempDir & TxtFile.Text
                                                intU = intU + 1
                                                GoTo JumpToPass
                                          End If
                                          'end 2020/03/25
                                    Else
                                          strErrCont = strErrCont & vbCrLf & "壓縮檔失敗:" & strNewName
                                          strNewName = ""
                                    End If
                                Else
                                    If UCase(strPass) = UCase("Thumbs.db") Then '刪除-瀏覽縮圖暫存檔
                                       If fs.FileExists(m_TempDir & strPass) Then
                                          Kill m_TempDir & strPass
                                       End If
                                       strNewName = ""
                                    Else
                                       strNewName = m_TempDir & strPass
                                    End If
                                End If
                                If strNewName <> "" Then '上傳檔案
                                    strExc(6) = "": strExc(2) = ""
                                    If strErrCont <> "" And InStr(strErrCont, strNewName) > 0 Then
                                        GoTo JumpToPass  '只顯示一次
                                    End If
                                    If PUB_ChkFileOpening(strNewName, , False) = True Then
                                        strErrCont = strErrCont & vbCrLf & "檔案正在使用中：" & strNewName
                                        GoTo JumpToPass
                                    End If
                                    If PUB_UploadCPFfile("0", strNewName, m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, nCP09, , , True, strExc(6), strExc(2)) = True Then
                                         If strExc(6) <> "" Then
                                             strErrCont = strErrCont & vbCrLf & strExc(6)
                                         End If
                                         iCount = iCount + 1
                                    Else
                                         strErrCont = strErrCont & vbCrLf & "上傳失敗：" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "(" & nCP09 & " ) : " & strExc(6)
                                    End If
                                    'Added by Lydia 2020/03/25 經過5筆資料,休眠1秒
                                    If intP Mod 4 = 0 Then
                                         Sleep 100
                                    End If
                                    'end 2020/03/25
                                End If
                                 intP = intP + 1
                            End If
JumpToPass:                         '計數：intP
                            strPass = Dir()
                         Loop
                         '上傳完後,直接刪除案號資料夾
                         If intU = 0 Then
                             intP = 0
                             strExc(2) = Dir(m_TempDir, vbDirectory)
                             Do While strExc(2) <> ""
                                  If strExc(2) <> "." And strExc(2) <> ".." Then
                                      If GetAttr(m_TempDir & strExc(2)) = vbDirectory Then
                                           intP = intP + 1
                                      Else
                                           intP = intP + 1
                                      End If
                                  End If
                                  strExc(2) = Dir()
                             Loop
                             If intP = 0 Then
                                 '刪除資料夾
                                 m_TempDir = Mid(m_TempDir, 1, Len(m_TempDir) - 1)
                                 Call PUB_KillTempFolder(Val(m_CP02), Mid(m_TempDir, 1, InStrRev(m_TempDir, "\") - 1))
                             End If
                         End If
                         '隔日,人工刪除第一層資料夾
JumpTo01:
                    Next intI
                End If
                
'專利案件分析：本所案號
                strKey = "\專利案件\"
                intK = InStr(UCase(strDefPath), strKey)
                If intK > 0 Then
                    m_CP01 = "FCP"
                    m_CP10 = cnt專利案件
                    m_CP03 = "0": m_CP04 = "00"
                    strExc(9) = ""
                    strExc(1) = Mid(strDefPath, intK + Len(strKey))
                    tmpArr1 = Split(strExc(1), "\") '用\區隔路徑層級
                    If Len(tmpArr1(0)) = 3 Then '前3碼相同:放同一層
                        strExc(9) = strExc(9) & "," & strDefPath
                    ElseIf Len(tmpArr1(0)) = 7 Then 'ex. 前3碼200_299: 底下再分前3碼子資料夾
                        '讀取:前3碼子資料夾
                        strPass = Dir(strDefPath, vbDirectory)
                        Do While strPass <> ""
                             If strPass <> "." And strPass <> ".." Then
                                 If GetAttr(strDefPath & strPass) = vbDirectory Then
                                      If Val(strPass) > 100 And Val(strPass) < 400 Then  'Added by Lydia 2020/03/23 判斷是否為前3碼本所案號
                                           strExc(9) = strExc(9) & "," & strDefPath & strPass
                                      End If 'Added by Lydia 2020/03/23
                                 End If
                             End If
                             strPass = Dir()
                        Loop
                    End If
                    '逐案號資料夾上傳
                    If strExc(9) <> "" Then
                        tmpArr1 = Empty
                        tmpArr1 = Split(Mid(strExc(9), 2), ",")
                        nMax = UBound(tmpArr1)
                        
                        For intI = 0 To nMax
                            m_TempDir = Trim(tmpArr1(intI))
                            If Right(m_TempDir, 1) <> "\" Then m_TempDir = m_TempDir & "\" '與抓子資料夾有關
                            
                            nCP09 = ""
                            intP = 0
                           '1.先讀取資料夾的所有檔案，拿掉檔名的Unicode字
'                            Set fs = CreateObject("Scripting.FileSystemObject") 'Remove by Lydia 2020/03/23
                            If Trim(m_TempDir) = "" Then GoTo JumpTo02
                            If Not fs.FolderExists(m_TempDir) Then
                                strErrCont = strErrCont & vbCrLf & "資料夾不存在：" & m_TempDir
                                GoTo JumpTo02
                            Else
                                'Set fs = CreateObject("Scripting.FileSystemObject") 'Remove by Lydia 2020/03/23
                                Set fso = fs.GetFolder(m_TempDir)
                                For Each fl In fso.files
                                   TxtFile.Text = fl.Name
                                   strNewName = TxtFile.Text 'Added by Lydia 2020/03/27
                                   strErr = "拿掉Unicode字:" & TxtFile.Text
                                    If TxtFile.Text <> fl.Name Then
                                        strNewName = Replace(TxtFile.Text, "?", "x")
                                        'Added by Lydia 2020/03/26 檢查是否有相同新檔名
                                        If ChkNewFileName(m_TempDir, strNewName) = False Then
                                             strErrCont = strErrCont & vbCrLf & "有相同新檔名:" & m_TempDir & TxtFile.Text & "，請確認檔名！"
                                             GoTo JumpTo02
                                        Else
                                        'end 2020/03/26
                                             fl.Name = strNewName
                                             TxtFile.Text = strNewName 'Added by Lydia 2020/03/27
                                        End If 'end 2020/03/26
                                        
                                    'Added by Lydia 2020/03/27 整合更名
                                    Else
                                         If ChkNewFileName(m_TempDir, strNewName) = False Then
                                             strErrCont = strErrCont & vbCrLf & "有相同新檔名:" & m_TempDir & TxtFile.Text & "，請確認檔名！"
                                             GoTo JumpTo02
                                         Else
                                              If strNewName <> TxtFile.Text Then
                                                   fl.Name = strNewName
                                              End If
                                         End If
                                    'end 2020/03/27
                                    End If
                                Next
                            End If
                             '2.上傳檔案，若有子資料夾則壓縮為.zip
JumpToReDir2:
                             strPass = Dir(m_TempDir, vbDirectory)
                             intP = 0: intU = 0
                             Do While strPass <> ""
                                If strPass <> "." And strPass <> ".." Then
                                    TxtFile.Text = strPass
                                    If InStr(TxtFile.Text, ".") = 0 And InStr(TxtFile.Text, "?") > 0 Then '子資料夾有Unicode
                                        strErrCont = strErrCont & vbCrLf & "子資料夾名稱有Unicode: " & m_TempDir & TxtFile.Text
                                        intU = intU + 1
                                        GoTo JumpToPass2
                                    End If
                                    strNewName = ""
                                    If GetAttr(m_TempDir & strPass) = vbDirectory Then
                                        strNewName = m_TempDir & Mid(strPass, 1, 20) & "." & Format(FileDateTime(m_TempDir & strPass), "YYYYMMDDHHMMSS")
                                        If ZipFolder(m_TempDir & strPass, strNewName) = True Then
                                              strNewName = strNewName & ".zip"
                                              Call PUB_KillTempFolder(strPass, m_TempDir)
                                              Sleep 100 'Added by Lydia 2020/03/25 休眠1秒
                                              'Added by Lydia 2020/03/25 判斷子資料夾是否存在
                                              If Not fs.FolderExists(m_TempDir & strPass) Then
                                                    intP = 0: intU = 0 '因為重讀資料夾,計數歸0
                                                    GoTo JumpToReDir2  '因為刪除子資料夾所以要重新查詢，不然子資料夾後的檔案無法讀到
                                              'Added by Lydia 2020/03/25
                                              Else
                                                     strErrCont = strErrCont & vbCrLf & "子資料夾未刪: " & m_TempDir & TxtFile.Text
                                                     intU = intU + 1
                                                    GoTo JumpToPass2
                                              End If
                                              'end 2020/03/25
                                        Else
                                              strErrCont = strErrCont & vbCrLf & "壓縮檔失敗: " & strNewName
                                              strNewName = ""
                                        End If
                                        
                                    Else
                                        If UCase(strPass) = UCase("Thumbs.db") Then '刪除-瀏覽縮圖暫存檔
                                           If fs.FileExists(m_TempDir & strPass) Then
                                              Kill m_TempDir & strPass
                                           End If
                                           strNewName = ""
                                        Else
                                           strNewName = m_TempDir & strPass
                                        End If
                                    End If
                                    If strNewName <> "" Then '上傳檔案
                                        strExc(1) = UCase(Mid(strNewName, InStrRev(strNewName, "\") + 1))
                                        If InStr(strExc(1), "FCP0") = 1 Then 'FCP開頭6碼
                                           m_CP02 = Mid(strExc(1), 4, 6)
                                        ElseIf InStr(strExc(1), "FCP") = 1 Then 'FCP開頭5碼
                                           m_CP02 = Format(Val(Mid(strExc(1), 4, 5)), "000000")
                                        Else
                                           m_CP02 = Format(Val(Mid(strExc(1), 1, 5)), "000000")
                                        End If
                                        strExc(6) = "": strExc(2) = ""
                                        If strErrCont <> "" And InStr(strErrCont, strNewName) > 0 Then
                                            GoTo JumpToPass2  '只顯示一次
                                        End If
                                        If PUB_ChkFileOpening(strNewName, , False) = True Then
                                            strErrCont = strErrCont & vbCrLf & "檔案正在使用中：" & strNewName
                                            GoTo JumpToPass2
                                        End If
                                        nCP09 = "" '因為專利案件不是一個案號一個資料夾,所以預設空白來抓案件是否有收文專利案件991
                                        If PUB_UploadCPFfile("0", strNewName, m_CP01, m_CP02, m_CP03, m_CP04, m_CP10, nCP09, , , , strExc(6), strExc(2)) = True Then
                                             If strExc(6) <> "" Then
                                                 strErrCont = strErrCont & vbCrLf & strExc(6)
                                             End If
                                             iCount = iCount + 1
                                        Else
                                             strErrCont = strErrCont & vbCrLf & "上傳失敗：" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "(" & nCP09 & " ) : " & strExc(6)
                                        End If
                                        'Added by Lydia 2020/03/25 經過5筆資料,休眠1秒
                                        If intP Mod 4 = 0 Then
                                             Sleep 100
                                        End If
                                        'end 2020/03/25
                                    End If
                                     intP = intP + 1
                                End If
JumpToPass2:                             '計數：intP
                                strPass = Dir()
                             Loop
                            '上傳完後,直接刪除案號資料夾
                            If intU = 0 Then
                                intP = 0
                                strExc(2) = Dir(m_TempDir, vbDirectory)
                                Do While strExc(2) <> ""
                                     If strExc(2) <> "." And strExc(2) <> ".." Then
                                         If GetAttr(m_TempDir & strExc(2)) = vbDirectory Then
                                              intP = intP + 1
                                         Else
                                              intP = intP + 1
                                         End If
                                     End If
                                     strExc(2) = Dir()
                                Loop
                                If intP = 0 Then
                                    '刪除資料夾
                                    m_TempDir = Mid(m_TempDir, 1, Len(m_TempDir) - 1)
                                    Call PUB_KillTempFolder(Mid(m_TempDir, InStrRev(m_TempDir, "\") + 1), Mid(m_TempDir, 1, InStrRev(m_TempDir, "\") - 1))
                                End If
                            End If
                            '隔日,人工刪除第一層資料夾
JumpTo02:
                        Next intI
                    End If
                End If
                cnnConnection.Execute " update Lydia_MoveList set r005='" & strSrvDate(1) & "' where r001='" & nPos & "' " '記錄已處理
                Sleep 100 'Added by Lydia 2020/03/25 休眠1秒
                .MoveNext
            Loop
        End With
        
        strCont = strCont & "~" & ChangeWStringToWDateString(strSrvDate(1)) & "(" & Format(ServerTime, "00:00:00") & ")"
        'Added by Lydia 2020/03/25 上班時間或下班時間有問題才發信
        'If ((strTime >= "2000" Or (strTime >= "0100" And strTime <= "0210")) And strErrCont <> "") Or (strTime >= "0800" And strTime <= "1600") Then
        If InStr(strSrvDate(1), Pub_GetSpecMan("XY搬檔指定日")) > 0 Or (strTime >= "2000" Or (strTime >= "0100" And strTime <= "0210")) Then
            PUB_SendMail strUserNum, "A3034", "", "外專利益衝突-搬檔作業完成", strCont & vbCrLf & "上傳成功：" & iCount & _
                                vbCrLf & vbCrLf & String(15, "=") & "錯誤記錄" & String(15, "=") & IIf(strErrCont <> "", strErrCont, ""), , , , , , , , , , , , , , False
        End If 'Added by Lydia 2020/03/25
    End If
                        
    Set rsAD = Nothing
    'Added by Lydia 2020/03/24
    Set rsB = Nothing
    Set fs = Nothing
    'end 2020/03/24
    
    bolMoveStatus = False
    lstHistory.AddItem Now & "--> 結束 外專利益衝突搬檔"
    
Exit Sub

ExceptCont:
     If strErr <> "" Or Err.Number <> 0 Then
         strErrCont = strErrCont & IIf(strErr <> "", vbCrLf & strErr & "---" & Err.Description, Err.Description)
         Resume Next
     End If
     
End Sub

'Added by Lydia 2020/03/19 將資料夾壓縮為.zip
Private Function ZipFolder(pFolder As String, pNewFolder As String) As Boolean
   Dim program_name As String, program_path As String
   Dim process_id As Long
   Dim process_handle As Long

   program_name = "C:\Program Files\7-Zip\7z.exe"
   '檢查執行檔
   If Dir(program_name) = "" Then
      MsgBox "未安裝 7-Zip 程式，壓縮檔產生失敗！。"
      Exit Function
   End If
    
On Error GoTo ShellError
        
   '不刪除舊檔,因為子資料夾有可能是壓縮檔解開來的,兩者皆要保留
   'If Dir(pFolder & ".zip") <> "" Then
   '   Kill pFolder & ".zip"
   'End If
   If Dir(pFolder & ".zip") <> "" Then
       pNewFolder = pNewFolder & ".New"
   End If
   
   '-y 指有相同檔案存在時, 直接覆蓋. 不給的話會需要在Console 給 yes/no. 適用於Automation
   '-p 解壓縮密碼
   'process_id = Shell("""" & program_name & """ a -pCTCB """ & pNewFolder & ".zip"" """ & pFolder & "\*""", vbNormalNoFocus)
   process_id = Shell("""" & program_name & """ a """ & pNewFolder & ".zip"" """ & pFolder & "\*""", vbNormalNoFocus)
    
    On Error GoTo 0

    ' Wait for the program to finish.
    ' Get the process handle.
    process_handle = OpenProcess(SYNCHRONIZE, 0, process_id)
    If process_handle <> 0 Then
        WaitForSingleObject process_handle, INFINITE
        CloseHandle process_handle
        ZipFolder = True
    End If
    Exit Function

ShellError:
    ZipFolder = False
    If Err.Number <> 0 Then
         '已在主程式寫Log
    End If
End Function

'Added by Lydia 2020/03/26 檢查是否有相同新檔名
Private Function ChkNewFileName(ByVal pPath As String, ByRef pNewName As String) As Boolean
Dim intRen As Integer
Dim m_FS
Dim strMidName As String

     If pPath = "" Or pNewName = "" Then Exit Function
          
     Set m_FS = CreateObject("Scripting.FileSystemObject")
     intRen = 1
     strMidName = pNewName

     'Added by Lydia 2020/03/27 整合更名: 特殊符號改為空白
     strMidName = Replace(strMidName, "'", " ") '單引號
     strMidName = Replace(strMidName, "#", " ")
     strMidName = Replace(strMidName, "|", " ")
     strMidName = Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(strMidName, "＃", " "), "◇", " "), "□", " "), "⊕", " "), "◎", " "), "*", " "), "△", " "), "＊", " "), "●", " "), "$", " "), "▲", " "), "◆", " ")
     strMidName = Replace(strMidName, "　", " ")
     strMidName = Replace(strMidName, "◎", " ")
     strMidName = Replace(strMidName, " .", ".")
     strMidName = Replace(strMidName, ". ", ".")
     strMidName = Replace(strMidName, "_.", ".")
     strMidName = Replace(strMidName, "._", ".")
     If strMidName = pNewName Then
        ChkNewFileName = True
        Exit Function
     End If
     'end 2020/03/27
     
     Do While intRen < 10
          If m_FS.FileExists(pPath & strMidName) Then
              strMidName = Mid(strMidName, 1, InStrRev(strMidName, ".") - 1) & "." & intRen & Mid(strMidName, InStrRev(strMidName, "."))
          Else
              ChkNewFileName = True
              Exit Do
          End If
          intRen = intRen + 1
     Loop
     If ChkNewFileName = True Then
         pNewName = strMidName
     End If
     
     Set m_FS = Nothing
End Function

'Add by Amy 2024/10/23
'判斷已執行多久
Private Function ChkRunTime(ByVal stChkTime As String, ByVal stEixtTime As String) As Boolean
   Dim stTime1 As String, stTime2 As String
   ChkRunTime = False
   stTime1 = ServerTime
   stTime2 = stChkTime
   '目前時間
   If Len(stTime1) = 5 Then
      stTime1 = Val(Left(stTime1, 1)) * 60 + Val(Mid(stTime1, 2, 2))
   Else
      stTime1 = Val(Mid(stTime1, 1, 2)) * 60 + Val(Mid(stTime1, 3, 2))
   End If
   '傳入要確認之時間
   If Len(stTime2) = 5 Then
      stTime2 = Val(Left(stTime2, 1)) * 60 + Val(Mid(stTime2, 2, 2))
   Else
      stTime2 = Val(Left(stTime2, 2)) * 60 + Val(Mid(stTime2, 3, 2))
   End If
   '已Run 超過 stEixtTime 分鐘
   If Val(stTime1) - Val(stTime2) >= Val(stEixtTime) Then
      ChkRunTime = True
   End If
End Function

'移檔
Private Function MoveFile(bolMoveBack As Boolean, stFormPath As String, stFileN As String, ByRef stMsg As String) As Boolean
   Dim oFilObj As FileSystemObject
   Dim stDate As String, stYear As String, stMon As String, stToPath As String, stHistoryPath As String
   
   MoveFile = False
   stMsg = ""
   stYear = Mid(strSrvDate(1), 1, 4)
   stMon = Mid(strSrvDate(1), 5, 2)
   stDate = Mid(strSrvDate(1), 7, 2)
   If bolMoveBack = True Then
      stToPath = Replace(stFormPath, "\History", "")
   Else
      stToPath = Replace(Replace(stFormPath, "\Success\", "\Success\History\"), "\Error\", "\Error\History\")
   End If
     
On Error GoTo ShowErr
    If Dir(stToPath, vbDirectory) = MsgText(601) Then
      MkDir Mid(stToPath, 1, Val(Len(stToPath)) - 1)
    End If
   If Dir(stToPath & stYear, vbDirectory) = MsgText(601) Then
      MkDir stToPath & stYear
   End If
   If Dir(stToPath & stYear & "\" & stMon, vbDirectory) = MsgText(601) Then
       MkDir stToPath & stYear & "\" & stMon
   End If
   If Dir(stToPath & stYear & "\" & stMon & "\" & stDate, vbDirectory) = MsgText(601) Then
      MkDir stToPath & stYear & "\" & stMon & "\" & stDate
   End If
  
   Set oFilObj = New FileSystemObject
   oFilObj.MoveFile stFormPath & stYear & "\" & stMon & "\" & stDate & "\" & stFileN, _
                                    stToPath & stYear & "\" & stMon & "\" & stDate & "\"
   Set oFilObj = Nothing
   If ChkXml(0, stToPath, stFileN) = False Then
      stMsg = "已執行移檔，但[" & stFileN & "]檔案未移至" & stToPath
      Exit Function
   End If
   MoveFile = True
   Exit Function
   
ShowErr:
   If Err.Number = 76 Then
      stMsg = "請確認資料夾[權限][共用]頁籤權限已開"
   Else
      stMsg = Err.Description
   End If
   stMsg = "移檔至" & stToPath & stYear & "\" & stMon & "\" & stDate & "\ 有誤-->" & stMsg
End Function

Private Function CheckIsRunning(pProcessName As String) As Boolean
   Dim Processes, Process
   Dim stOwner As String
      
   Set Processes = Interaction.GetObject("winmgmts:").ExecQuery("select * from Win32_Process where name='" & pProcessName & "'")
   If UCase(pProcessName) = UCase(App.EXEName & ".exe") Then
      If Processes.Count > 1 Then CheckIsRunning = True
   Else
      If Processes.Count > 0 Then CheckIsRunning = True
   End If
End Function
'end 2024/10/23

