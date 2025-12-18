VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frm040334 
   BorderStyle     =   1  '單線固定
   Caption         =   "證書PDF列印"
   ClientHeight    =   5565
   ClientLeft      =   2790
   ClientTop       =   3945
   ClientWidth     =   6420
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   6420
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4725
      Style           =   2  '單純下拉式
      TabIndex        =   28
      Top             =   5160
      Width           =   1545
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "啟動 PDF"
      Height          =   400
      Left            =   4725
      TabIndex        =   27
      Top             =   4710
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.TextBox txtFirstAdd 
      Height          =   315
      Left            =   4050
      TabIndex        =   23
      Text            =   "3"
      Top             =   4740
      Width           =   555
   End
   Begin VB.TextBox txtMaxSec 
      Height          =   315
      Left            =   4050
      TabIndex        =   19
      Text            =   "45"
      Top             =   5100
      Width           =   555
   End
   Begin VB.TextBox txtMinSec 
      Height          =   315
      Left            =   1470
      TabIndex        =   18
      Text            =   "5"
      Top             =   5100
      Width           =   555
   End
   Begin VB.TextBox txtByte 
      Height          =   315
      Left            =   1470
      TabIndex        =   17
      Text            =   "30000"
      Top             =   4740
      Width           =   975
   End
   Begin VB.TextBox txtPath2 
      Height          =   315
      Left            =   1860
      TabIndex        =   7
      Text            =   "\\pat3\GAZETTE\PXml"
      Top             =   1170
      Visible         =   0   'False
      Width           =   4395
   End
   Begin VB.FileListBox File2 
      Height          =   270
      Left            =   2490
      TabIndex        =   13
      Top             =   90
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.TextBox txtPDFPath 
      Height          =   315
      Left            =   1860
      TabIndex        =   8
      Text            =   "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      Top             =   1890
      Width           =   4395
   End
   Begin VB.ComboBox cmbPrinter2 
      Height          =   300
      Left            =   1860
      TabIndex        =   9
      Top             =   1530
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   4
      Left            =   2880
      MaxLength       =   2
      TabIndex        =   6
      Top             =   828
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   3
      Left            =   2628
      MaxLength       =   1
      TabIndex        =   5
      Top             =   828
      Width           =   255
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   2
      Left            =   1788
      MaxLength       =   6
      TabIndex        =   4
      Top             =   828
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   264
      Index           =   1
      Left            =   1308
      MaxLength       =   3
      TabIndex        =   3
      Text            =   "P"
      Top             =   828
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   264
      Index           =   0
      Left            =   1308
      MaxLength       =   7
      TabIndex        =   1
      Top             =   492
      Width           =   1335
   End
   Begin VB.OptionButton Option1 
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   60
      TabIndex        =   2
      Top             =   864
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "公告日："
      Height          =   180
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   540
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4785
      TabIndex        =   12
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3990
      TabIndex        =   11
      Top             =   60
      Width           =   756
   End
   Begin VB.ListBox List1 
      Height          =   1680
      ItemData        =   "frm040334.frx":0000
      Left            =   90
      List            =   "frm040334.frx":0007
      TabIndex        =   10
      Top             =   2970
      Width           =   6195
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   285
      Left            =   90
      TabIndex        =   25
      Top             =   2310
      Visible         =   0   'False
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Label lblProgress 
      Alignment       =   2  '置中對齊
      Caption         =   "( 0/0 )"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   135
      TabIndex        =   26
      Top             =   2610
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "第1個檔多加幾秒："
      Height          =   180
      Left            =   2490
      TabIndex        =   24
      Top             =   4740
      Width           =   1530
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "1個檔最多幾秒："
      Height          =   180
      Left            =   2670
      TabIndex        =   22
      Top             =   5100
      Width           =   1350
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "1個檔至少幾秒："
      Height          =   180
      Left            =   90
      TabIndex        =   21
      Top             =   5100
      Width           =   1350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "幾Byte算1秒："
      Height          =   180
      Left            =   300
      TabIndex        =   20
      Top             =   4740
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "公報PDF的存放路徑："
      Height          =   180
      Left            =   90
      TabIndex        =   16
      Top             =   1260
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "PDF執行檔路徑："
      Height          =   180
      Left            =   90
      TabIndex        =   15
      Top             =   1950
      Width           =   1560
   End
   Begin VB.Label Label6 
      Caption         =   "列印公報PDF印表機："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   1590
      Width           =   1755
   End
End
Attribute VB_Name = "frm040334"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Add By Sindy 2011/12/27
Option Explicit

Dim intWhere As Integer
Dim strReceiveNo As String '本所案號
Dim strTPB04 As String, strTPB05 As String
Dim i As Integer, j As Integer
'Modify By Sindy 2014/9/3
Dim m_DefaultPrinter As String
'Dim m_DefaultPrinter2 As String
Dim strPrinter As String
'2014/9/3 END
'Dim SeekPrint As Integer
Dim strTime As String
Dim m_AttachPath As String 'Added by Morgan 2021/6/25 公報PDF暫存路徑

Private Sub cmdOK_Click(Index As Integer)
Dim strTmp As String, rsTemp1 As New ADODB.Recordset, rsTemp2 As New ADODB.Recordset
Dim stET03 As String
Dim int_Copys As Integer
   
    
    Select Case Index
    Case 0 '確定
      cmdok(Index).Enabled = False 'Added by Morgan 2012/1/11
         'Add By Sindy 2011/12/27
         List1.Clear
'         If cmbPrinter2.ListIndex >= 0 Then
'             Set Printer = Printers(cmbPrinter2.ListIndex)
''             Printer.EndDoc
'         End If
         '2011/12/27 End
         'Modify By Sindy 2014/9/3
         '設定控制台預設印表機
         PUB_SetOsDefaultPrinter cmbPrinter2
         '系統印表機
         PUB_RestorePrinter cmbPrinter2
         '2014/9/3 END
         
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/18 清除查詢印表記錄檔欄位
        '公告日
        If Option1(0).Value = True Then
            If Text1(0).Text <> "" Then
                If Not ChkDate(Text1(0).Text) Then
                    Text1(0).SetFocus
                    TextInverse Text1(0)
                    Exit Sub
                End If
            Else
                MsgBox "公告日不得空白，請重新輸入 !", vbCritical
                Text1(0).SetFocus
                Exit Sub
            End If
            
            'Add By Sindy 2011/12/27
            'Removed by Morgan 2021/6/25 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
            'If GetFilePath(DBDATE(Text1(0))) = False Then
            '   Me.txtPath2.SetFocus
            '   Exit Sub
            'End If
            'end 2021/6/25
            '2011/12/27 End
            
            Screen.MousePointer = vbHourglass
            pub_QL05 = pub_QL05 & ";" & Option1(0).Caption & Text1(0) 'Add By Sindy 2010/11/18
            
            'Modified by Morgan 20121/1
            strExc(0) = "select pa01,pa02,pa03,pa04,pa11,pa14,cp09,cp66,cp67 from patent,caseprogress " & _
                        "Where pa14=" & DBDATE(Text1(0)) & "  and pa01='P' and pa09='000' " & _
                        "and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 " & _
                        "and cp10='1603' " & _
                        "order by cp66 asc,cp67 asc,cp09 asc "
            intI = 1
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With rsTemp2
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/18
                  Do While Not .EOF
                     Call GetPDFCopys(.Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04"), "" & .Fields("PA11"), int_Copys)
                     .MoveNext
                  Loop
               End With
               
               If List1.ListCount > 0 Then
                  'Modified by Morgan 2012/1/16
                  'Call PrintPDF
                  Call PrinBatchPdf
                  MsgBox "列印完成 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
               Else
                  MsgBox "列印完成 !", vbInformation
               End If
            Else
                InsertQueryLog (0) 'Add By Sindy 2010/11/18
                MsgBox "無符合條件之資料 !", vbInformation
            End If
            Screen.MousePointer = vbDefault
        '本所案號
        Else
            If Text1(2) = "" Then
                MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
                Text1(2).SetFocus
                Exit Sub
            End If
            strTmp = Text1(1) & Text1(2)
            If Text1(3).Text = "" Then
                strTmp = strTmp & "0"
            Else
                strTmp = strTmp & Text1(3).Text
            End If
            If Text1(4).Text = "" Then
                strTmp = strTmp & "00"
            Else
                strTmp = strTmp & Text1(4).Text
            End If
            Screen.MousePointer = vbHourglass
            pub_QL05 = pub_QL05 & ";" & Option1(1).Caption & Text1(1) & "-" & Text1(2) & "-" & IIf(Text1(3) = "", "0", Text1(3)) & "-" & IIf(Text1(4) = "", "00", Text1(4)) 'Add By Sindy 2010/11/18
            
            strExc(0) = "select pa01,pa02,pa03,pa04,pa11,pa14 from patent,caseprogress " & _
                        "Where " & ChgPatent(strTmp) & " " & _
                        "and pa01=cp01 and pa02=cp02 and pa03=cp03 and pa04=cp04 " & _
                        "and cp10='1603' " & _
                        "order by cp66 asc,cp67 asc "
            intI = 1
            Set rsTemp2 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               With rsTemp2
                  InsertQueryLog (.RecordCount) 'Add By Sindy 2010/11/18
                  Do While Not .EOF
                     If "" & .Fields("PA14") > "" Then
                        'Modified by Morgan 2021/6/25 公報PDF改抓卷宗區，不再往pat3讀取避免當機沒開的情形
                        'If GetFilePath("" & .Fields("PA14")) = False Then
                        '   Me.txtPath2.SetFocus
                        If PUB_GetGazettePDF(.Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04")) = False Then
                           MsgBox .Fields(0) & "案公告公報卷宗區PDF檔讀取失敗！", vbExclamation
                        'end 2021/6/25
                           Screen.MousePointer = vbDefault
                           Exit Sub
                        End If
                     Else
                        MsgBox "此本所案號無公告日 !", vbInformation
                        Screen.MousePointer = vbDefault
                        Text1(2).SetFocus
                        Exit Sub
                     End If
                     
                     Call GetPDFCopys(.Fields("PA01"), .Fields("PA02"), .Fields("PA03"), .Fields("PA04"), "" & .Fields("PA11"), int_Copys)
                     .MoveNext
                  Loop
               End With
               
               If List1.ListCount > 0 Then
                  'Modified by Morgan 2012/1/16
                  'Call PrintPDF
                  Call PrinBatchPdf
                  MsgBox "列印完成 ! (列印PDF花費時間：" & strTime & "  " & time() & ")", vbInformation
               Else
                  MsgBox "列印完成 !", vbInformation
               End If
            Else
               InsertQueryLog (0) 'Add By Sindy 2010/11/18
               MsgBox "無符合條件之資料 !", vbInformation
            End If
            Screen.MousePointer = vbDefault
        End If
        
        cmdok(Index).Enabled = True 'Added by Morgan 2012/1/11
        
         'Modify By Sindy 2014/9/3
         '還原控制台預設印表機
         PUB_SetOsDefaultPrinter strPrinter
         '還原系統中預設印表機
         PUB_RestorePrinter m_DefaultPrinter
         '2014/9/3 END
    Case 1 '結束
        Unload Me
    End Select
End Sub

Private Sub GetPDFCopys(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, StrPA11 As String, ByRef int_Copys As Integer)
Dim strFileName As String
   
'Modified by Morgan 2012/1/17 改都只要印一份
'   int_Copys = 0
'
'   '由員工檔取得列印份數 (北部的員工印2份, 其它地區的員工印3份)
'   strExc(0) = "SELECT ST06 FROM STAFF WHERE ST01='" & PUB_GetAKindSalesNo(strPA01, strPA02, strPA03, strPA04) & "' "
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   int_Copys = 3
'   If intI = 1 Then
'      If RsTemp.Fields(0).Value = "1" Then
'         int_Copys = 2
'      Else
'         int_Copys = 3
'      End If
'   End If
int_Copys = 1
'end 2012/1/17
   
   'Modify By Sindy 2013/1/4
   'strFileName = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\" & StrPA11 & "-P01.pdf"
   'Modified by Morgan 2021/6/25 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
   'strFileName = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05 & "\" & StrPA11 & ".pdf"
   If PUB_GetGazettePDF(strPA01, strPA02, strPA03, strPA04, True, m_AttachPath, strFileName) = False Then
      strFileName = ""
   End If
   'end 2021/6/25
   '2013/1/4 End
   
   If strFileName <> "" Then
      List1.AddItem strFileName & " " & int_Copys
   End If
End Sub

Private Function GetFilePath(strDate As String) As Boolean
On Error GoTo ErrHnd
   
   GetFilePath = True
   Exit Function 'Added by Morgan 2021/6/25 公報改抓卷宗區，不再往pat3讀取避免當機沒開的情形
   
   If IsEmptyText(txtPath2) = True Then
      MsgBox "請輸入公報PDF的存放路徑！", vbExclamation + vbOKOnly
      GetFilePath = False
      Exit Function
   End If
   If Right(Trim(txtPath2), 1) = "\" Then txtPath2 = Left(txtPath2, Len(txtPath2) - 1)
   
   strTPB04 = Format(Val(Val(Left(strDate, 4)) - 1911) - 62, "00")
   j = Val(Mid(strDate, 5, 2))
   i = (j - 1) * 3
   j = Val(Right(strDate, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   ElseIf j >= 21 Then
      i = i + 3
   End If
   strTPB05 = Format(i, "00")
   
   File2.path = txtPath2 & "\img_1\isu0" & strTPB04 & "0" & strTPB05
   File2.Refresh
   If File2.ListCount = 0 Then
      MsgBox "公報PDF的存放路徑中無" & strTPB04 & "卷" & strTPB05 & "期資料！"
      GetFilePath = False
      Exit Function
   End If
   
   Exit Function
   
ErrHnd:
   If Err.NUMBER = 76 Then
      MsgBox "公報PDF的存放路徑中無" & strTPB04 & "卷" & strTPB05 & "期資料！"
   Else
      MsgBox Err.Description, vbCritical
   End If
   GetFilePath = False
End Function

'Removed by Morgan 2017/9/30 不再使用
''Add by Morgan 2010/2/3
''設定作業系統預設印表機
'Public Sub SetOsDefaultPrinter(strPrinter As String)
'   Dim idx As Integer, strOsPrinter As String
'   If strPrinter <> "" Then
'      strOsPrinter = PUB_GetOsDefaultPrinter
'      If strPrinter <> strOsPrinter Then
'         '檢查有存在才做
'         For idx = 0 To Printers.Count - 1
'            If Printers(idx).DeviceName = strPrinter Then
'               Printer.TrackDefault = True 'Modify by Sindy 2011/12/30
'               CreateObject("WScript.Network").SetDefaultPrinter strPrinter
'               Exit For
'            End If
'         Next
'      End If
'   End If
'End Sub

'Removed by Morgan 2017/5/15 不再使用
'Private Sub PrintPDF()
'Dim i As Integer, k As Integer, strTemp As Variant
'Dim RetVal, intFileCnt As Integer
'Dim ff1 As Integer
'Dim MySize, dblSec As Double, dblCntSec As Double
'
'   strTime = time()
'   intFileCnt = 0
'
''   '設定控制台預設印表機
''   If cmbPrinter2.ListIndex >= 0 Then
''      PUB_SetOsDefaultPrinter Printers(cmbPrinter2.ListIndex).DeviceName
''   End If
'
'   If ff1 > 0 Then Close #ff1
'   ff1 = FreeFile
'   Open txtPath2 & "\專利證書" & strTPB04 & "卷" & strTPB05 & "期" & "列印PDF時間資訊.txt" For Output As #ff1
'
'   For i = 0 To List1.ListCount - 1
'      strTemp = Split(List1.List(i), " ")
'
'      For k = 0 To Val(strTemp(1)) - 1 '列印份數
'         intFileCnt = intFileCnt + 1
'   '      AcroPDF1.src = List1.List(i)
'   '      AcroPDF1.LoadFile (List1.List(i))
'   '      DoEvents
'   '      Sleep 5000
'   '      AcroPDF1.printAll
'   '      DoEvents
'   '      Sleep 5000
'         'C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe
'         'RetVal = SHELL(txtPDFPath & " /p /h /t " & List1.List(i), vbHide) '/p /h /t
'         RetVal = SHELL(txtPDFPath & " /p /h " & strTemp(0), vbHide) '/p /h /t
'         MySize = FileLen(strTemp(0))   '傳回檔案長度 (以 Byte 為單位)
'         'DoEvents
'         'If i = 0 And k = 0 Then
''         If CDbl(MySize) <= CDbl(512000) Then
''            Sleep 5000
''         ElseIf CDbl(MySize) > CDbl(512000) And CDbl(MySize) < CDbl(1048576) Then
''            Sleep 8000
''         Else
''            Sleep 10000
''         End If
'         '依檔案大小決定秒數
'         If Val(txtByte) = 0 Then
'            MsgBox "[幾Byte算1秒]此欄位不可空白 !", vbInformation
'            txtByte.SetFocus
'            Close #ff1
'            Exit Sub
'         End If
'         dblCntSec = Round(MySize / CDbl(txtByte), 0)
'         If dblCntSec <= 0 Then
'            dblSec = (CDbl(txtMinSec) * 1000)
'         Else
'            dblSec = (dblCntSec * 1000)
'         End If
'         '開第1個檔案時多加幾秒
'         If Val(txtFirstAdd) > 0 Then
'            If i = 0 And k = 0 Then dblSec = dblSec + (CDbl(txtFirstAdd) * 1000)
'         End If
'         '至少幾秒
'         If Val(txtMinSec) > 0 Then
'            If dblSec < (CDbl(txtMinSec) * 1000) Then
'               dblSec = (CDbl(txtMinSec) * 1000)
'            End If
'         End If
'         '最多幾秒
'         If Val(txtMaxSec) > 0 Then
'            If dblSec > (CDbl(txtMaxSec) * 1000) Then
'               dblSec = (CDbl(txtMaxSec) * 1000)
'            End If
'         End If
'         Sleep dblSec
'
'         If k = 0 Then
'            Print #ff1, Left(i + 1 & "     ", 5) & List1.List(i) & " " & MySize & " " & dblCntSec & " " & dblSec
'         End If
'      Next k
'
'   Next i
'
''   '還原控制台預設印表機
''   If cmbPrinter2.ListIndex >= 0 Then
''      PUB_SetOsDefaultPrinter m_DefaultPrinter
''   End If
'
'   Print #ff1, "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
'   Print #ff1, "列印時間：" & strTime & "  " & time()
'   Print #ff1, "檔案數量：" & intFileCnt
'   Close #ff1
'
'End Sub

Private Sub cmdRun_Click()
   SHELL txtPDFPath, Combo1.ItemData(Combo1.ListIndex)
End Sub

Private Sub Form_Load()
'Dim SeekPrintL As Integer
'Dim i As Integer, j As Integer
   
   MoveFormToCenter Me
   intWhere = 國外_FC
   
'   If Pub_StrUserSt03 = "M51" Then
'      m_DefaultPrinter = Printer.DeviceName
'   Else
'      m_DefaultPrinter = "ApeosPort-II 6000(10F影印機.機密雙面)"
'   End If
'   m_DefaultPrinter2 = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
'   For i = 0 To Printers.Count - 1
'      Set Printer = Printers(i)
'      cmbPrinter2.AddItem Printer.DeviceName, j
'      j = j + 1
'      If Printer.DeviceName = m_DefaultPrinter Then
'         SeekPrint = i
'      End If
'   Next i
'   Set Printer = Printers(SeekPrint)
''   Printer.EndDoc
'   cmbPrinter2.Text = cmbPrinter2.List(SeekPrint)
'   If m_DefaultPrinter <> m_DefaultPrinter2 Then
'      SetOsDefaultPrinter Printers(SeekPrint).DeviceName
'   End If
   'Modify By Sindy 2014/9/3
   PUB_SetPrinter Me.Name, cmbPrinter2, m_DefaultPrinter
   strPrinter = PUB_GetOsDefaultPrinter '抓控制台目前預設的印表機
   '2014/9/3 END
   
   List1.Clear
   
   'Modified by Morgan 2012/1/13 改用 API 抓
   If Pub_StrUserSt03 = "M51" Then
      cmdRun.Visible = True
      Combo1.AddItem "vbHide", 0
      Combo1.ItemData(0) = 0
      Combo1.AddItem "vbMinimizedNoFocus", 0
      Combo1.ItemData(0) = 6
      Combo1.ListIndex = 0
      
   '   txtPDFPath = "C:\Program Files\Adobe\Reader 8.0\Reader\AcroRd32.exe"
      'Me.Height = 6120 'Removed by Morgan 2012/1/11
   Else
   '   'C:\Program Files\Adobe\Acrobat 8.0\Acrobat\Acrobat.exe
   '   'C:\Program Files\Adobe\Acrobat 7.0\Reader\AcroRd32.exe
   '   txtPDFPath = "C:\Program Files\Adobe\Acrobat 8.0\Acrobat\Acrobat.exe"
   '   'Modified by Morgan 2012/1/11
   '   'Me.Height = 2700
   '   Me.Height = 2925
      Me.Height = 3255
   End If
   'Modify By Sindy 2014/9/3
   'SetFileAssociation
   txtPDFPath = PUB_SetFileAssociation
   '2014/9/3 END
   'end 2012/1/13
   
   Option1_Click 0
   
   'Added by Morgan 2021/6/25 公報PDF暫存路徑
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   'end 2021/6/25
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   If m_DefaultPrinter <> m_DefaultPrinter2 Then
'      PUB_SetOsDefaultPrinter m_DefaultPrinter2
'   End If
   'Modify By Sindy 2014/9/3
   '若印表機變動, 則更新列印設定
   If Me.cmbPrinter2.Text <> Me.cmbPrinter2.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.cmbPrinter2.Name, "0", "0", Me.cmbPrinter2.Text
   End If
   '2014/9/3 END
   
   Set frm040334 = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
 Dim txt As TextBox, i As Integer
On Error Resume Next
   For Each txt In Text1
      txt.Enabled = False
   Next
   Select Case Index
      Case 0
         Text1(0).Enabled = True
         Text1(0).SetFocus
      Case 1
         For i = 2 To 4
            Text1(i).Enabled = True
         Next
         Text1(2).SetFocus
   End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
   If Text1(Index) = "" Then Exit Sub
   If Option1(0).Value = True Then
      If Index = 0 Then
         If Text1(Index).Text <> "" Then
            If Not ChkDate(Text1(Index).Text) Then
               Text1(Index).SetFocus
               TextInverse Text1(Index)
            End If
         Else
            MsgBox "公告日不得空白，請重新輸入 !", vbCritical
            Text1(Index).SetFocus
         End If
      End If
   Else
      If Index = 1 Then
         If Text1(Index).Text = "" Then
            MsgBox "本所案號不得空白，請重新輸入 !", vbCritical
            Text1(Index).SetFocus
         End If
      End If
   End If
End Sub

'Added by Morgan 2012/1/13
Private Sub PrinBatchPdf()
    
Dim program_name As String
Dim process_id As Long
Dim process_handle As Long
Dim ii As Integer, kk As Integer
'Modified by Morgan 2021/6/25
'Dim strTemp As Variant
Dim strTemp(1) As String
'end 2021/6/25
Dim ff1 As Integer
Dim strPrinterName As String
Dim intFileCnt As Integer
Dim MySize
   
   strTime = time()
   
   ProgressBar1.max = List1.ListCount
   ProgressBar1.Value = 0
   lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
   ProgressBar1.Visible = True
   lblProgress.Visible = True
   DoEvents

   program_name = txtPDFPath
   strPrinterName = cmbPrinter2

    ' Start the program.
On Error GoTo ShellError
    
    '因為第 2 個以後開啟的 Reader 才會印完後自動關閉,所以固定先開一個空的程式,全部印完後再關閉
    'Modified by Morgan 2017/5/15 路徑可能含空白,改加雙引號
    process_id = SHELL("""" & program_name & """", vbHide)
    process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
    
On Error GoTo 0
    
   If ff1 > 0 Then Close #ff1
   ff1 = FreeFile
   'Modified by Morgan 2021/6/25
   'Open txtPath2 & "\專利證書" & strTPB04 & "卷" & strTPB05 & "期" & "列印PDF時間資訊.txt" For Output As #ff1
   Open m_AttachPath & "\專利證書" & strTPB04 & "卷" & strTPB05 & "期" & "列印PDF時間資訊.txt" For Output As #ff1
   'end 2021/6/25
   
   For ii = 0 To List1.ListCount - 1
      'Modified by Morgan 2021/6/25
      'strTemp = Split(List1.List(ii), " ")
      intI = InStrRev(List1.List(ii), " ")
      strTemp(0) = Left(List1.List(ii), intI - 1)
      strTemp(1) = Mid(List1.List(ii), intI + 1)
      'end 2021/6/25
      For kk = 1 To Val(strTemp(1)) '列印份數
         intFileCnt = intFileCnt + 1
         mdiMain.tmrConnect.Tag = 0
         'Modified by Morgan 2017/5/15 改呼叫共用函數
         'PrintOnePdf program_name, " /n /t """ & strTemp(0) & """ """ & strPrinterName & """"
         PUB_PrintOnePdf program_name, " /n /t """ & strTemp(0) & """ """ & strPrinterName & """"
      Next
      
      ProgressBar1.Value = ProgressBar1.Value + 1
      lblProgress = "( " & ProgressBar1.Value & "/" & ProgressBar1.max & " )"
      DoEvents
      MySize = FileLen(strTemp(0))
      Print #ff1, Left(ii + 1 & "     ", 5) & List1.List(ii) & " " & MySize
   Next
    
   TerminateProcess process_handle, 0&
   CloseHandle process_handle
   
   Print #ff1, "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Print #ff1, "列印時間：" & strTime & "  " & time()
   Print #ff1, "檔案數量：" & intFileCnt
   Close #ff1
   
   ProgressBar1.Visible = False
   lblProgress.Visible = False
   Exit Sub

ShellError:
    MsgBox " " & _
        program_name & vbCrLf & _
        Err.Description, vbOKOnly Or vbExclamation, _
        "Error"
End Sub


