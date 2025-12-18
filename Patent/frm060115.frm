VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060115 
   BorderStyle     =   1  '單線固定
   Caption         =   "電子收據匯入卷宗區"
   ClientHeight    =   5760
   ClientLeft      =   40
   ClientTop       =   280
   ClientWidth     =   8980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8980
   Begin VB.CommandButton cmdPrint 
      Caption         =   "產生文字檔"
      Height          =   400
      Left            =   3180
      TabIndex        =   13
      Top             =   90
      Width           =   1065
   End
   Begin VB.CommandButton cmdPath 
      Height          =   330
      Left            =   8460
      Picture         =   "frm060115.frx":0000
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   630
      Width           =   350
   End
   Begin VB.ListBox List1 
      Height          =   3820
      ItemData        =   "frm060115.frx":0102
      Left            =   960
      List            =   "frm060115.frx":0104
      TabIndex        =   1
      Top             =   1020
      Width           =   7845
   End
   Begin VB.Frame Frame1 
      Height          =   465
      Left            =   30
      TabIndex        =   7
      Top             =   5250
      Width           =   8895
      Begin VB.TextBox Text2 
         Alignment       =   2  '置中對齊
         BackColor       =   &H00FF0000&
         Height          =   300
         Left            =   30
         TabIndex        =   9
         Top             =   120
         Width           =   8820
      End
   End
   Begin VB.FileListBox File1 
      Height          =   420
      Left            =   1560
      TabIndex        =   6
      Top             =   60
      Visible         =   0   'False
      Width           =   735
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   405
      Left            =   1020
      TabIndex        =   5
      Top             =   60
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   864
      _ExtentY        =   723
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frm060115.frx":0106
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "C:\E-Set\elecReceiptDir"
      Top             =   660
      Width           =   7455
   End
   Begin VB.CommandButton cmdTransFile 
      Caption         =   "執行(&E)"
      Height          =   400
      Left            =   6000
      TabIndex        =   2
      Top             =   90
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6990
      TabIndex        =   3
      Top             =   90
      Width           =   912
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
      Height          =   2310
      Left            =   450
      TabIndex        =   12
      Top             =   5910
      Width           =   2625
      _ExtentX        =   4639
      _ExtentY        =   4092
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "匯入結果："
      Height          =   180
      Left            =   60
      TabIndex        =   10
      Top             =   1050
      Width           =   900
   End
   Begin VB.Label Label2 
      Caption         =   "轉檔中, 請稍候 . . ."
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   15.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   30
      TabIndex        =   8
      Top             =   4950
      Width           =   8895
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "CSV檔："
      Height          =   180
      Left            =   60
      TabIndex        =   4
      Top             =   720
      Width           =   690
   End
End
Attribute VB_Name = "frm060115"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/26 Form2.0已檢查 (無需修改的物件)
'Create By Sindy 2015/8/17 請作單:1080103-01
Option Explicit

Dim strCText1 As String, strEText2 As String
Dim TmpData As String
Dim arrData As Variant
Dim int_i As Integer

Private Declare Function SendMessageByNum Lib "user32" _
  Alias "SendMessageA" (ByVal hWnd As Long, ByVal _
  wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Const LB_SETHORIZONTALEXTENT = &H194

Dim oFileSys As New FileSystemObject
Dim oFile As File
Dim bolReadCP As Boolean


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPath_Click()
   Dim stPath As String
   
   With cd1
   .Filter = "Supported files|*.csv"
   .FilterIndex = 0
   If txtPath = "" Then
      .InitDir = PUB_Getdesktop
   Else
      stPath = Left(txtPath, InStrRev(txtPath, "\") - 1)
      If PUB_ChkDir(stPath) = True Then
         .InitDir = stPath
      Else
         .InitDir = PUB_Getdesktop
      End If
   End If
   .ShowOpen
   If Trim(.FileName) <> "" Then
      SaveSetting "TAIE", "FCPReceipt", UCase(Me.Name) & "Dir", .FileName
      txtPath.Text = .FileName
   End If
   End With
End Sub

Private Function ReadTextFile(Optional pCharset As String = "big5") As String
   Dim adoStream As ADODB.Stream
   Dim var_String As Variant
   Dim strFileName As String
   
   strFileName = txtPath
   Set adoStream = New ADODB.Stream
   'adoStream.Charset = "UTF-8"
   adoStream.Charset = pCharset
   adoStream.Open
   adoStream.LoadFromFile strFileName
   ReadTextFile = adoStream.ReadText
   adoStream.Close
   Set adoStream = Nothing
End Function

'去除多餘的符號
Private Function GetStr(ByVal pContent As String) As String
   pContent = Trim(pContent)
   '去除右邊逗號
   If Right(pContent, 1) = "," Then
      pContent = Left(pContent, Len(pContent) - 1)
   End If
   '去除左邊雙引號
   Do While Left(pContent, 1) = """"
      pContent = Mid(pContent, 2)
   Loop
   '去除右邊雙引號
   Do While Right(pContent, 1) = """"
      pContent = Left(pContent, Len(pContent) - 1)
   Loop
   GetStr = pContent
End Function

Private Function Import2DB(Optional pImpRecs As Integer, Optional pSkipRecs As Integer) As Boolean
   Dim strText As String
   Dim arrRow() As String
   Dim arrCell() As String
   Dim idx1 As Integer, idx2 As Integer
   Dim stSQL As String, stValues As String, intR As Integer
   Dim stER01 As String, stER07 As String, stER21 As String, stER05 As String, stER03 As String, stER06 As String, stER08 As String
   Dim bolIsTaie As Boolean
   Dim stFolder As String, stFileName As String
   Dim arrERidx(22) As Integer
   Dim adoRst As New ADODB.Recordset
   Dim arrColNames() As String
   Dim strNewCol As String
   Dim iRecs As Integer
   
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim strCP09 As String
   Dim strCP10 As String
   Dim stReName As String
   Dim strAttachPath As String
   Dim strErrMsg As String
   Dim strTotRow As String
   Dim dblMaxWidth As Double
   Dim bolManyCnt As Boolean, errTxt As String
   Dim ii As Integer
   Dim intK As Integer
   Dim strCP84 As String
   
   Const cDelimiter As String = """,""" '欄位區隔符號
   
   pImpRecs = 0
   pSkipRecs = 0
   
On Error GoTo ErrHnd
   
   strText = ReadTextFile
   strText = Replace(strText, Chr(9), "")
   strText = Replace(strText, Chr(13) & Chr(10), Chr(10))
   arrRow = Split(strText, Chr(10))
   arrCell = Split(arrRow(LBound(arrRow)), cDelimiter)
   ReDim arrColNames(UBound(arrCell)) As String
   
   strNewCol = ""
   arrERidx(15) = -1
   For idx1 = LBound(arrCell) To UBound(arrCell)
      strText = GetStr(arrCell(idx1))
      Select Case strText
      Case "收據號碼"
         arrColNames(idx1) = "ER01": arrERidx(1) = idx1
         
      Case "案件種類"
         arrColNames(idx1) = "ER02": arrERidx(2) = idx1
      
      Case "金額"
         arrColNames(idx1) = "ER03": arrERidx(3) = idx1
         
      Case "繳費時間"
         arrColNames(idx1) = "ER04": arrERidx(4) = idx1
      
      Case "收發文號"
         arrColNames(idx1) = "ER05": arrERidx(5) = idx1
         
      Case "開立日期"
         arrColNames(idx1) = "ER06": arrERidx(6) = idx1
         
      Case "案號"
         arrColNames(idx1) = "ER07": arrERidx(7) = idx1
         
      Case "費用類別"
         arrColNames(idx1) = "ER08": arrERidx(8) = idx1
         
      Case "案件名稱"
         arrColNames(idx1) = "ER09": arrERidx(9) = idx1
         
      Case "繳費方式"
         arrColNames(idx1) = "ER10": arrERidx(10) = idx1
         
      Case "申請人"
         arrColNames(idx1) = "ER11": arrERidx(11) = idx1
         
      Case "年度"
         arrColNames(idx1) = "ER12": arrERidx(12) = idx1
      
      Case "專利證書號"
         arrColNames(idx1) = "ER13": arrERidx(13) = idx1
         
      Case "商標註冊號"
         arrColNames(idx1) = "ER14": arrERidx(14) = idx1
         
      Case "檔案名稱"
         arrColNames(idx1) = "ER15": arrERidx(15) = idx1
         
      Case "繳款人"
         arrColNames(idx1) = "ER19": arrERidx(19) = idx1
      
      Case "實際繳款人" 'Added by Morgan 2017/5/5
         arrColNames(idx1) = "ER20": arrERidx(20) = idx1
      
      Case "自訂案件編號" 'Added by Morgan 2017/9/26
         arrColNames(idx1) = "ER21": arrERidx(21) = idx1
         
      Case "收據種類" 'Added by Morgan 2017/9/26
         arrColNames(idx1) = "ER22": arrERidx(22) = idx1
      'Added by Morgan 2018/2/23 目前沒用,略過
      'Modfied by Morgan 2024/2/2 +案由資訊(E-Set下載的CSV檔欄位名稱不同)
      Case "案由", "案由資訊"
         
      Case Else
         strNewCol = strNewCol & strText & vbCrLf
      End Select
   Next
   
   'Removed by Morgan 2024/2/2 沒有匯入收據檔不必彈訊息
   'If strNewCol <> "" Then
   '   If MsgBox("有新增欄位如下將不會匯入，是否確定要繼續？" & vbCrLf & vbCrLf & strNewCol, vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
   '      Exit Function
   '   End If
   'End If
   'end 2024/2/2
   
   If UBound(arrRow) = LBound(arrRow) Then
      MsgBox "無資料!!" & vbCrLf & vbCrLf & arrCell(0), vbCritical, "匯入失敗"
      Exit Function
   End If
   
   Screen.MousePointer = vbHourglass
   strTotRow = File1.ListCount
   Me.Height = 6120
   dblMaxWidth = 8820
   Text2.Width = 0
   
   For idx1 = LBound(arrRow) + 1 To UBound(arrRow)
      Text2.Width = dblMaxWidth / Val(strTotRow) * (idx1): DoEvents
      
      If arrRow(idx1) <> "" Then
         arrCell = Split(arrRow(idx1), cDelimiter)
         
         stFolder = ""
         stValues = ""
         stER01 = GetStr(arrCell(arrERidx(1))) '收據號碼
         stER03 = GetStr(arrCell(arrERidx(3))) '金額 Add By Sindy 2019/5/15
         stER05 = GetStr(arrCell(arrERidx(5))) '收發文號 Add By Sindy 2019/3/29
         stER06 = GetStr(arrCell(arrERidx(6))) '開立日期 Add By Sindy 2019/5/15
         stER07 = GetStr(arrCell(arrERidx(7))) '案號
         stER08 = GetStr(arrCell(arrERidx(8))) '費用類別
         stER21 = GetStr(arrCell(arrERidx(21))) '自訂案件編號
         errTxt = "": bolManyCnt = False
         
         'Add By Sindy 2019/5/15
         '讀取與開立日期相符的自動扣款日,FCP案進度資料
         If bolReadCP = False Then
            'Modify By Sindy 2019/5/22 原抓開立日期,因退費的開立日期會是之前日期,所以改抓系統日
            '   cp152=" & stER06 ==> cp152=" & strSrvDate(1)
            strExc(0) = "select '',cp01||'-'||cp02||'-'||cp03||'-'||cp04,cp09,cp84,cp10,cp64" & _
               " from caseprogress" & _
               " where cp152=" & strSrvDate(1) & " and cp01='FCP'" & _
               " and cp158>0 and cp159=0" & _
               " order by cp158 asc,cp82 asc"
            intI = 0
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If RsTemp.RecordCount > 0 Then
                  Set Grid1.Recordset = RsTemp
                  GridHead
                  bolReadCP = True
                  Me.Tag = stER06
               Else
                  Exit Function
               End If
            Else
               Exit Function
            End If
         End If
         '2019/5/15 END
         
         '電子檔資料
         stFolder = Left(txtPath, InStrRev(txtPath, "\"))
         If Right(stFolder, 1) = "\" Then stFolder = Left(stFolder, Len(stFolder) - 1)
         stFileName = GetStr(arrCell(arrERidx(15)))
         strErrMsg = ""
         '以申請案號判斷是那一個本所案號
         bolIsTaie = False
         If stER07 <> "" Then
            strExc(0) = "select * from patent where pa11='" & stER07 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               bolIsTaie = True
               strCP01 = RsTemp.Fields("pa01")
               strCP02 = RsTemp.Fields("pa02")
               strCP03 = RsTemp.Fields("pa03")
               strCP04 = RsTemp.Fields("pa04")
            End If
         End If
         If bolIsTaie Then
            If Dir(stFolder & "\" & stFileName) <> "" Then
               If strCP01 = "FCP" Then
                  'Add By Sindy 2019/3/29 發生同案號有二道進度同時來收據 ex:FCP-054464
                  '                       stER05=收發文號
                  '有比對到智慧局收文文號資料者:
                  'Modify By Sindy 2025/9/9 進度檔同一智慧局收文文號, 則為這張收據的資料明細
'                  strExc(0) = "select count(*),sum(cp84) from caseprogress" & _
'                              " where cp01='" & strCP01 & "'" & _
'                              " and cp02='" & strCP02 & "'" & _
'                              " and cp03='" & strCP03 & "'" & _
'                              " and cp04='" & strCP04 & "'" & _
'                              " and cp158>0 and cp159=0" & _
'                              " and cp84>0 and instr(cp64,'" & stER05 & "')>0" & _
'                              " order by cp27 desc"
                  strExc(0) = "select count(*),sum(nvl(cp84,0)) from caseprogress" & _
                              " where cp01='" & strCP01 & "'" & _
                              " and cp02='" & strCP02 & "'" & _
                              " and cp03='" & strCP03 & "'" & _
                              " and cp04='" & strCP04 & "'" & _
                              " and cp158>0 and cp159=0" & _
                              " and instr(cp64,'" & stER05 & "')>0" & _
                              " order by cp152 asc" 'order by cp27 desc
                  '2025/9/9 END
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     'Add By Sindy 2019/5/23
                     If RsTemp.Fields(0) > 0 Then '有資料
                     '2019/5/23 END
                        If Val(RsTemp.Fields(0)) > 1 Then
                           bolManyCnt = True
                        Else
                           bolManyCnt = False
                        End If
                        strCP84 = "" & RsTemp.Fields(1) '金額
                        
                        '記錄文號資料
                        If bolManyCnt = True Then
                           'Modify By Sindy 2025/9/30 多筆的,有時有掛金額有掛日期,不會是在想歸卷的文號裡,所以日期先檢查
                           '   +,cp152
                           strExc(0) = "select cp09,cp152 from caseprogress" & _
                                       " where cp01='" & strCP01 & "'" & _
                                       " and cp02='" & strCP02 & "'" & _
                                       " and cp03='" & strCP03 & "'" & _
                                       " and cp04='" & strCP04 & "'" & _
                                       " and cp158>0 and cp159=0" & _
                                       " and cp84>0 and instr(cp64,'" & stER05 & "')>0" & _
                                       " order by cp27 desc"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              RsTemp.MoveFirst
                              
                              'Add By Sindy 2025/9/30
                              '檢查日期是否相符
                              If Val(stER06) <> Val("" & RsTemp.Fields("cp152")) Then
                                 errTxt = strCP01 & "-" & strCP02 & "," & stER01 & "扣款日期不符"
                              End If
                              '2025/9/30 END
                              
                              Do While Not RsTemp.EOF
                                 Call FindCPData(RsTemp.Fields("cp09"), strCP01, strCP02, strCP03, strCP04)
                                 RsTemp.MoveNext
                              Loop
                           End If
                        End If
                        
                        'Add By Sindy 2025/9/30
                        If errTxt = "" Then
                        '2025/9/30 END
      '                     If RsTemp.Fields(0) = 1 Then '只查詢到一筆才歸
                           '檢查金額是否相符
                           If Val(stER03) = Val(strCP84) Then
                              'Modify By Sindy 2020/4/8 增加 .data.pdf 比對資料
                              'ex:FCP-063017 分割 / 續行母案再審
                              If bolManyCnt = True Then
                                 'Modify By Sindy 2025/9/9 取消 and cp84>0, ex:主動修正不會掛發文規費但有申請書 FCP-073604,FCP-073692
                                 strExc(0) = "select distinct cp01,cp02,cp03,cp04,cp09,cp10,cp152,cp84,cp27 from caseprogress,casepaperpdf" & _
                                             " where cp01='" & strCP01 & "'" & _
                                             " and cp02='" & strCP02 & "'" & _
                                             " and cp03='" & strCP03 & "'" & _
                                             " and cp04='" & strCP04 & "'" & _
                                             " and cp158>0 and cp159=0" & _
                                             " and instr(cp64,'" & stER05 & "')>0" & _
                                             " and cp09=cpp01(+) and instr(upper(cpp02),upper('.data.pdf'))>0" & _
                                             " order by cp152 asc" 'order by cp27 desc
                              Else
                              '2020/4/8 END
                                 strExc(0) = "select distinct cp01,cp02,cp03,cp04,cp09,cp10,cp152,cp84,cp27 from caseprogress" & _
                                             " where cp01='" & strCP01 & "'" & _
                                             " and cp02='" & strCP02 & "'" & _
                                             " and cp03='" & strCP03 & "'" & _
                                             " and cp04='" & strCP04 & "'" & _
                                             " and cp158>0 and cp159=0" & _
                                             " and cp84>0 and instr(cp64,'" & stER05 & "')>0" & _
                                             " order by cp27 desc"
                              End If
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           Else
                              errTxt = strCP01 & "-" & strCP02 & "," & stER01 & "金額不符"
                           End If
                        End If
                     Else
                        intI = 0
                     End If
                  End If
                  '2019/3/29 END
                  If errTxt = "" Then
                     'If intI = 0 Or bolManyCnt = True Then
                     If intI = 0 Or RsTemp.RecordCount > 1 Then
                        'Modify By Sindy 2019/5/15
                        intK = 0
                        If stER08 = "年費" Then '費用類別
                           '再檢查是否有年費,若有,則歸最近一道年費進度
                           strExc(0) = "select distinct cp01,cp02,cp03,cp04,cp09,cp10,cp152,cp84,cp27 from caseprogress" & _
                                       " where cp01='" & strCP01 & "'" & _
                                       " and cp02='" & strCP02 & "'" & _
                                       " and cp03='" & strCP03 & "'" & _
                                       " and cp04='" & strCP04 & "'" & _
                                       " and cp158>0 and cp159=0" & _
                                       " and cp84>0 and cp10='" & 年費 & "'" & _
                                       " order by cp27 desc"
                           intK = 1
                           Set RsTemp = ClsLawReadRstMsg(intK, strExc(0))
                           intI = intK
                        End If
                        If intK = 0 Then
                        '2019/5/15 END
                           '歸最後發文有發文規費及申請書的文號
                           strExc(0) = "select distinct cp01,cp02,cp03,cp04,cp09,cp10,cp152,cp84,cp27 from caseprogress,casepaperpdf" & _
                                       " where cp01='" & strCP01 & "'" & _
                                       " and cp02='" & strCP02 & "'" & _
                                       " and cp03='" & strCP03 & "'" & _
                                       " and cp04='" & strCP04 & "'" & _
                                       " and cp158>0 and cp159=0" & _
                                       " and cp84>0 and cp09=cpp01(+) and instr(upper(cpp02),upper('.data.pdf'))>0" & _
                                       " order by cp27 desc"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        End If
                     End If
                     If intI = 1 Then
                        If bolManyCnt = False Then '多筆金額和日期在前面比對
                           '檢查金額是否相符
                           If Val(stER03) <> Val("" & RsTemp.Fields("cp84")) Then
                              errTxt = strCP01 & "-" & strCP02 & "," & stER01 & "金額不符"
                           '檢查日期是否相符
                           ElseIf Val(stER06) <> Val("" & RsTemp.Fields("cp152")) Then
                              errTxt = strCP01 & "-" & strCP02 & "," & stER01 & "扣款日期不符"
                           End If
                        End If
                        
                        '記錄案號資料
                        Call FindCPData("", strCP01, strCP02, strCP03, strCP04)
                        
                        If errTxt <> "" Then
                           pSkipRecs = pSkipRecs + 1
                           List1.AddItem stFolder & "\" & stFileName & " => " & errTxt
                        Else
                           '匯入卷宗區:開始
                           strCP09 = RsTemp.Fields("cp09")
                           strCP10 = RsTemp.Fields("cp10")
                           stReName = strCP01 & strCP02 & _
                                      IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & _
                                      "." & strCP10 & ".RECEIPT.pdf"
                           strExc(0) = "select count(*) from casepaperpdf" & _
                                       " where cpp01='" & strCP09 & "'" & _
                                       " and instr(upper(cpp02),upper('.RECEIPT.'))>0 and substr(upper(cpp02),-4)='.PDF'"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              If RsTemp.Fields(0) > 0 Then
                                 stReName = strCP01 & strCP02 & _
                                            IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & _
                                            "." & strCP10 & ".RECEIPT." & Val(RsTemp.Fields(0)) & ".pdf"
                              End If
                           End If
                           
                           'Add By Sindy 2020/9/17 收據號碼:109DP099413;存入進度備註
                           'Modified by Morgan 2022/10/27 CP64後面要加空白檢查否則NULL時不會更新
                           strSql = "update caseprogress set cp64=cp64||'收據號碼:" & stER01 & ";'" & _
                                    " where cp09='" & strCP09 & "'" & _
                                    " and instr(cp64||' ','收據號碼:" & stER01 & ";')=0"
                           cnnConnection.Execute strSql, intI
                           '2020/9/17 END
                           
                           '上傳檔案至卷宗區
                           Set oFile = oFileSys.GetFile(stFolder & "\" & stFileName)
                           If SaveAttFile_PDF(strCP09, stFolder & "\" & stFileName, stReName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , "Y", True, , , , strErrMsg) = True Then
                              '上傳個案
                              If Pub_StrUserSt03 = "M51" Then
                                 strAttachPath = PUB_Getdesktop & "\FCP"
                              Else
                                 'Modified by Lydia 2024/07/22 改用變數
                                 'strAttachPath = "\\typing2\FCP_workflow\FCP\" & Left(strCP02, 3)
                                 strAttachPath = "\\" & strTyping2Path & "\FCP_workflow\FCP\" & Left(strCP02, 3)
                              End If
                              If Dir(strAttachPath, vbDirectory) = "" Then
                                 MkDir strAttachPath
                              End If
                              If Pub_StrUserSt03 = "M51" Then
                                 strAttachPath = PUB_Getdesktop & "\FCP\" & strCP01 & strCP02
                              Else
                                 'Modified by Lydia 2024/07/22 改用變數
                                 'strAttachPath = "\\typing2\FCP_workflow\FCP\" & Left(strCP02, 3) & "\" & strCP01 & strCP02
                                 strAttachPath = "\\" & strTyping2Path & "\FCP_workflow\FCP\" & Left(strCP02, 3) & "\" & strCP01 & strCP02
                              End If
                              If Dir(strAttachPath, vbDirectory) = "" Then
                                 MkDir strAttachPath
                              End If
                              
                              stReName = strCP01 & strCP02 & _
                                         IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & _
                                         ".RECEIPT." & strSrvDate(1) & ".pdf"
                              If Dir(strAttachPath & "\" & stReName) = "" Then
                                 FileCopy stFolder & "\" & stFileName, strAttachPath & "\" & stReName
                              Else
                                 For int_i = 1 To 15
                                    stReName = strCP01 & strCP02 & _
                                               IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & _
                                               ".RECEIPT." & int_i & "." & strSrvDate(1) & ".pdf"
                                    If Dir(strAttachPath & "\" & stReName) = "" Then
                                       FileCopy stFolder & "\" & stFileName, strAttachPath & "\" & stReName
                                       Exit For
                                    End If
                                 Next int_i
                              End If
                              
                              '刪除檔案
                              Kill stFolder & "\" & stFileName
                              pImpRecs = pImpRecs + 1
                              List1.AddItem stFolder & "\" & stFileName & " => " & strAttachPath & "\" & stReName
                           Else
                              pSkipRecs = pSkipRecs + 1
                              List1.AddItem stFolder & "\" & stFileName & " => " & strCP01 & Val(strCP02) & _
                                                                       IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & _
                                                                       " 存檔失敗" & strErrMsg
                           End If
                        End If
                     Else
                        pSkipRecs = pSkipRecs + 1
                        List1.AddItem stFolder & "\" & stFileName & " => " & strCP01 & Val(strCP02) & _
                                                                    IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & _
                                                                    " 找不到可以歸卷的文號"
                     End If
                  Else
                     pSkipRecs = pSkipRecs + 1
                     List1.AddItem stFolder & "\" & stFileName & " => " & errTxt
                  End If
               ElseIf strCP01 = "P" Then
                  '刪除檔案
                  Kill stFolder & "\" & stFileName
                  pImpRecs = pImpRecs + 1
                  List1.AddItem stFolder & "\" & stFileName & " => P案件，電子檔已刪除"
               Else
                  pSkipRecs = pSkipRecs + 1
                  List1.AddItem stFolder & "\" & stFileName & " => 非FCP,P案件"
               End If
            End If
         Else
            pSkipRecs = pSkipRecs + 1
            If stER07 = "" Then
               List1.AddItem stFolder & "\" & stFileName & " => 案號空白"
            Else
               List1.AddItem stFolder & "\" & stFileName & " => 非台一案件"
            End If
         End If
      End If
   Next
   
   '檢查系統有當天扣款,但無收據
   If Grid1.Rows > 1 And Grid1.TextMatrix(1, 1) <> "" Then '有資料才比對
      For ii = 1 To Grid1.Rows - 1
         If Grid1.TextMatrix(ii, 0) = "" Then
            pSkipRecs = pSkipRecs + 1
            List1.AddItem "系統有當天扣款,但無收據 => " & Grid1.TextMatrix(ii, 1) & " , " & Grid1.TextMatrix(ii, 2)
         End If
      Next ii
   End If
   
   Import2DB = True
   Set adoRst = Nothing
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & vbCrLf & "收據號碼:" & stER01, vbCritical, "匯入失敗"
   End If
   Set adoRst = Nothing
   
End Function

Private Sub FindCPData(strCP09 As String, strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String)
Dim ii As Integer
   
   If Grid1.Rows > 1 And Grid1.TextMatrix(1, 1) <> "" Then '有資料才比對
      For ii = 1 To Grid1.Rows - 1
         If Grid1.TextMatrix(ii, 0) = "" Then
            If strCP09 <> "" Then
               If Grid1.TextMatrix(ii, 2) = strCP09 Then
                  Grid1.TextMatrix(ii, 0) = "V"
                  Exit For
               End If
            Else
               If Grid1.TextMatrix(ii, 1) = strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 Then
                  Grid1.TextMatrix(ii, 0) = "V"
                  Exit For
               End If
            End If
         End If
      Next ii
   End If
End Sub

Private Sub cmdTransFile_Click()
Dim strTit As String, strMsg As String, nResponse
'Dim dblFCnt As Double
'Dim rsTmp As New ADODB.Recordset
Dim strTime As String ', strTotRow As String
Dim TmpPath As String
'Dim m_PA01 As String, m_PA02 As String, m_PA03 As String, m_PA04 As String, m_CP09 As String
'Dim stFileName As String, stReName As String
'Dim strCP10 As String
'Dim fs, f
   
On Error GoTo ErrHand
   
   strTime = time()
   List1.Clear
   '檢查資料
   If IsEmptyText(txtPath) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入電子檔路徑！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtPath.SetFocus
      Exit Sub
   End If
   TmpPath = txtPath
   If Dir(TmpPath) <> "" Then
      If InStrRev(txtPath, "\") > 0 Then
         TmpPath = Left(txtPath, InStrRev(txtPath, "\") - 1)
      End If
      If Dir(TmpPath) <> "" Then
         strTit = "檢核資料"
         strMsg = "無此電子檔路徑！"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtPath.SetFocus
         Exit Sub
      End If
   End If
   File1.path = TmpPath
   File1.Refresh
   If File1.ListCount = 0 Then
      strTit = "檢核資料"
      strMsg = "此資料夾無資料！"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      txtPath.SetFocus
      Exit Sub
   End If
   'MsgBox "請先選擇CSV檔案!", vbExclamation
   
   'Add by Sindy 2019/5/15 清空Grid1資料
   Grid1.Clear
   Grid1.Rows = 2
   GridHead
   '2019/5/15 END
   
   Dim iImpRecs As Integer, iSkipRecs As Integer
   Call Import2DB(iImpRecs, iSkipRecs)
   
   If iSkipRecs > 0 Then
      MsgBox "匯入完畢，尚有未處理的電子檔，請查看！（" & iSkipRecs & "筆）", vbInformation
   Else
      MsgBox "匯入完畢！", vbInformation
   End If
   
   Call SetListScroll(List1)
   Screen.MousePointer = vbDefault
   Me.Height = 5310
   
   Exit Sub
   
ErrHand:
   Call SetListScroll(List1)
   Screen.MousePointer = vbDefault
   ErrorMsg
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

Private Sub Form_Load()
   MoveFormToCenter Me
   
   '讀取前次設定路徑
   txtPath = GetSetting("TAIE", "FCPReceipt", UCase(Me.Name) & "Dir", "")
   If txtPath <> "" Then
      strExc(1) = Left(txtPath, InStrRev(txtPath, "\"))
      strExc(0) = Dir(strExc(1) & "*.CSV")
      If strExc(0) <> "" Then
         txtPath = strExc(1) & strExc(0)
      End If
   End If
   
   GridHead 'Add By Sindy 2019/5/15
   Me.Height = 5310
End Sub

'Add By Sindy 2019/5/15
Private Sub GridHead()
Dim i As Integer
   
   FixGrid Grid1
   With Grid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1200: .Text = "本所案號"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1400: .Text = "發文規費"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1400: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1400: .Text = "進度備註"
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
   bolReadCP = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060115 = Nothing
End Sub

Private Sub txtPath_GotFocus()
   InverseTextBox txtPath
End Sub

'產生文字檔
Private Sub cmdPrint_Click()
Dim ii As Integer
Dim ff1 As Integer
Dim strFileName As String
   
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   'Modify By Sindy 2019/5/22 原抓開立日期,因退費的開立日期會是之前日期,所以改抓系統日
   'strFileName = Me.Tag & "電子收據匯入卷宗區錯誤清單.txt"
   strFileName = strSrvDate(1) & "電子收據匯入卷宗區錯誤清單.txt"
   Open PUB_Getdesktop & "\" & strFileName For Output As ff1
   Print #ff1, "備註：改字型Fixedsys標準11號字以橫式上下左右各10MM列印"
   Print #ff1, "匯入結果"
   Print #ff1, "================================================================================================"
   For ii = 0 To List1.ListCount
      Print #ff1, List1.List(ii)
   Next ii
   Close ff1
   MsgBox "產生完畢！請至下列位置列印清單：" & PUB_Getdesktop & "\" & strFileName, vbInformation
End Sub
