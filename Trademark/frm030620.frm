VERSION 5.00
Begin VB.Form frm030620 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標公報大陸清單列印"
   ClientHeight    =   3600
   ClientLeft      =   2790
   ClientTop       =   3950
   ClientWidth     =   6170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   6170
   Begin VB.FileListBox File2 
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Index           =   0
      Left            =   1530
      TabIndex        =   0
      Text            =   "C:\temp\XmlTrans"
      Top             =   960
      Width           =   3675
   End
   Begin VB.CheckBox Check2 
      Caption         =   "公告明細 (Word檔)"
      Height          =   255
      Left            =   1530
      TabIndex        =   3
      Top             =   2010
      Width           =   4575
   End
   Begin VB.CheckBox Check1 
      Caption         =   "清單"
      Height          =   255
      Left            =   1530
      TabIndex        =   2
      Top             =   1710
      Value           =   1  '核取
      Width           =   2085
   End
   Begin VB.TextBox txtTBD01 
      BackColor       =   &H8000000F&
      Enabled         =   0   'False
      Height          =   264
      Left            =   1530
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1350
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   756
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4125
      TabIndex        =   4
      Top             =   120
      Width           =   756
   End
   Begin VB.Label Label7 
      Caption         =   "(               筆)"
      Height          =   210
      Left            =   2670
      TabIndex        =   13
      Top             =   1380
      Width           =   1230
   End
   Begin VB.Label Label6 
      Caption         =   " (檔案位置名稱c:\temp\商標公報xx卷xx期大陸公告明細)"
      Height          =   255
      Left            =   1740
      TabIndex        =   12
      Top             =   2310
      Width           =   4365
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "注意：當程式正在產生公告明細時，請暫時不要使用Word"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   60
      TabIndex        =   11
      Top             =   2820
      Width           =   6060
   End
   Begin VB.Label Label5 
      Caption         =   "並且電腦不可以設定螢幕保護裝置"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   780
      TabIndex        =   10
      Top             =   3090
      Width           =   3615
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "商標圖檔案路徑："
      Height          =   180
      Left            =   60
      TabIndex        =   8
      Top             =   1020
      Width           =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "報表種類："
      Height          =   210
      Left            =   600
      TabIndex        =   7
      Top             =   1710
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期："
      Height          =   210
      Left            =   600
      TabIndex        =   6
      Top             =   1380
      Width           =   900
   End
End
Attribute VB_Name = "frm030620"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim PLeft(1 To 5) As Integer
Dim strTemp(1 To 5) As String
Dim iPgae As Integer, iLine As Integer
Dim QueryCntRow As Integer

'加入代表圖用
Const msoBringInFrontOfText = 4
Const msoFalse = 0
Const msoLineSolid = 1
Const msoLineSingle = 1
Const msoTrue = -1
Const msoPictureAutomatic = 1

Dim m_WordFilePath As String
Dim m_intFileCnt As Integer
Dim m_AppAddr As String '商標註冊人地址
Dim m_AppName As String '商標註冊人
Dim m_AppAddrZip As String '商標註冊人地址郵遞區號
Dim bolRetry As Boolean '是否已發生錯誤且重試
Dim m_iRow As Integer
Dim m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String
Dim custarea As String   '業務區
Dim custsales As String  '智權人員


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If Trim(txtTBD01) = "" Then
            MsgBox "公報卷期不可空白！", vbInformation, "輸入錯誤！"
            txtTBD01.SetFocus
            Exit Sub
         End If
         If Check1.Value = 0 And Check2.Value = 0 Then
            MsgBox "報表種類至少選一項！", vbExclamation
            Check1.SetFocus
            Exit Sub
         End If
         If Check2.Value = 1 And txtPath(0).Text = "" Then
            MsgBox "檔案路徑不可空白！", vbExclamation
            txtPath(0).SetFocus
            Exit Sub
         End If
         
         QueryCntRow = 0
         If Check1.Value = 1 Then
            StrMenu
            If QueryCntRow = 0 Then Exit Sub
         End If
         If Check2.Value = 1 Then
            Process
            If QueryCntRow = 0 Then Exit Sub
         End If
         
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()
Dim i As Integer

Printer.Orientation = 1 '1.直印 2.橫印

strSql = "select TMBM01,TM05,TM06,TM07,TMBM05,TMBM06,TMBM08 " & _
         "from TMBulletin,TMBulletindata,trademark " & _
         "where TMBM07='" & Val(txtTBD01) & "' " & _
         "and TMBM01=TBD02 and TMBM02=TBD03 and TBD15='B' and TBD16='1' " & _
         "and TMBM01=TM15(+) " & _
         "order by TMBM06,TMBM01 asc "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   Screen.MousePointer = vbHourglass
   QueryCntRow = m_rs.RecordCount
   With m_rs
      m_rs.MoveFirst
      iLine = 0
      Do While Not m_rs.EOF
         For i = 1 To 5
            strTemp(i) = ""
         Next i
         strTemp(1) = CheckStr(m_rs.Fields("TMBM01"))
         strTemp(2) = Left(CheckStr(m_rs.Fields("TM05")) & "                              ", 30)
         strTemp(3) = CheckStr(m_rs.Fields("TMBM05"))
         strTemp(4) = CheckStr(m_rs.Fields("TMBM06"))
         strTemp(5) = CheckStr(m_rs.Fields("TMBM08"))
         
         If iLine > 53 Or iLine = 0 Then
            If iLine <> 0 Then
               Printer.NewPage
            End If
            iLine = 1
            PrintTitle
         End If
         PrintDetail '列印表中
         m_rs.MoveNext
      Loop
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iLine * 300
      Printer.Print "共計 " & m_rs.RecordCount & " 筆"
   End With
Else
   MsgBox "查詢無資料！", vbExclamation + vbOKOnly
   Exit Sub
End If
Printer.EndDoc
Screen.MousePointer = vbDefault
If Check2.Value = 0 Then
   MsgBox "列印完畢！", vbExclamation + vbOKOnly
End If
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("商標公報" & Left(txtTBD01, Len(txtTBD01) - 2) & "卷" & Right(txtTBD01, 2) & "期大陸清單") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "商標公報" & Left(txtTBD01, Len(txtTBD01) - 2) & "卷" & Right(txtTBD01, 2) & "期大陸清單"

Printer.Font.Size = 12
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "審定號數"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "商標名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "地區名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "代理人名稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "商品類別"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 1800
PLeft(3) = 6500
PLeft(4) = 8000
PLeft(5) = 9500
End Sub

Sub PrintDetail()
Dim i As Integer
   For i = 1 To 5
      Printer.CurrentX = PLeft(i)
      Printer.CurrentY = iLine * 300
      Printer.Print strTemp(i)
   Next i
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   If QueryData = False Then
      cmdOK(0).Enabled = False
   Else
      cmdOK(0).Enabled = True
   End If
   
   If Pub_StrUserSt03 = "M51" Then
      txtPath(0).Enabled = True
   Else
      txtPath(0).Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030620 = Nothing
End Sub

Private Sub txtPath_GotFocus(Index As Integer)
   TextInverse txtPath(Index)
End Sub

Private Sub txtTBD01_GotFocus()
   InverseTextBox txtTBD01
End Sub

' 公報卷期
Private Sub txtTBD01_Validate(Cancel As Boolean)
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   Cancel = False
   If IsEmptyText(txtTBD01) = False Then
      If IsNumeric(txtTBD01) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公報卷期只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTBD01_GotFocus
         Exit Sub
      ElseIf Val(Right(Me.txtTBD01.Text, 2)) < 1 Or Val(Right(Me.txtTBD01.Text, 2)) > 24 Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "公報期數輸入錯誤!!!"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTBD01_GotFocus
         Exit Sub
      End If
   End If
End Sub

Private Function Process() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strTime As String
Dim fs As Object
Dim i As Integer
   
   On Error GoTo ErrHnd
   
   Process = False
   
   strSql = "select tmbm06 " & _
              "From tmbulletinowner,tmbulletindata,tmbulletin " & _
             "Where tbor02=1 " & _
               "and tbor01=tbd02 and tbor06=tbd03 " & _
               "and tbd02=tmbm01(+) and tbd03=tmbm02(+) and tbd15='B' and TBD16='1' " & _
            "group by tmbm06 " & _
            "order by tmbm06 asc "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      QueryCntRow = rsTmp.RecordCount
      
      strTime = time()
      If Right(txtPath(0), 1) = "\" Then txtPath(0) = Left(txtPath(0), Len(txtPath(0)) - 1)
      File2.path = txtPath(0).Text & "\imagesdata"
      File2.Refresh
      If File2.ListCount = 0 Then
         MsgBox "找不到商標圖檔！"
         Exit Function
      End If
      
      m_WordFilePath = "c:\temp"
      
'      Set fs = CreateObject("Scripting.FileSystemObject")
'      fs.DeleteFolder m_WordFilePath, True
'NotFolder76:
'      fs.CreateFolder m_WordFilePath
      
      '產生Word檔
      m_intFileCnt = 0
      bolRetry = True
      Screen.MousePointer = vbHourglass
      cnnConnection.BeginTrans
      
      rsTmp.MoveFirst
      For m_iRow = 1 To rsTmp.RecordCount
         ' 列印定稿
         If WordEdit(rsTmp.Fields("tmbm06")) = False Then
            GoTo ErrHnd
         End If
         
'         If (m_iRow Mod 20) = 0 Or m_iRow = rsTmp.RecordCount Then
'            g_WordAp.Documents.Save
'            g_WordAp.Documents.Close
'            bolRetry = True
'         End If
'         If (m_iRow Mod 40) = 0 Then
'              Exit For
'         End If
         rsTmp.MoveNext
      Next m_iRow
   Else
      MsgBox "查詢無資料！", vbExclamation + vbOKOnly
      Exit Function
   End If
   
   If bolRetry = False Then
      g_WordAp.Documents.Save
      g_WordAp.Documents.Close
'      g_WordAp.Visible = True
'      g_WordAp.WindowState = wdWindowStateMaximize
   End If
   
   cnnConnection.CommitTrans
   Screen.MousePointer = vbDefault
   
   Process = True
   Set rsTmp = Nothing
   Set g_WordAp = Nothing
   MsgBox "作業完成！Word檔案產生在" & m_WordFilePath & "下。（花費時間：" & strTime & "  " & time() & "）"
   Exit Function
   
ErrHnd:
'   If Err.Number = 76 Then
'      GoTo NotFolder76
'   Else
   If Err.Number <> 0 Then
      'cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   If Err.Number <> 70 Then '70.沒有使用權限
      cnnConnection.RollbackTrans
   End If
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
   Set g_WordAp = Nothing
   'Resume
End Function

Private Function WordEdit(strKey1 As String) As Boolean
   'Add by Morgan 2011/10/26 +信頭
   Dim stFileName As String
   Dim iPicNo As Integer
   Dim iPicNo2 As Integer
   Dim oShape
   'Added by Morgan 2020/3/30
   If strSrvDate(1) >= 智慧所更名日 Then
      PUB_GetLetterPicID "1", "T", iPicNo, iPicNo2
   Else
   'end 2020/3/30
      iPicNo = 12
      iPicNo2 = 11
   End If 'Added by Morgan 2020/3/30
   'end 2011/10/26
   Dim rsTmp As New ADODB.Recordset
   Dim i As Integer, j As Integer
   Dim strTBD01 As String, strTBD01_2 As String
   Dim strTBD02 As String
   Dim strTBD03 As String, strTBD03_2 As String
   Dim strTBD04 As String
   Dim strTBD05 As String
   Dim strTBD06 As String
   Dim strTBD07 As String
   Dim strTBD08 As String
   Dim strTBD09 As String
   Dim strTBD10 As String
   Dim strTBD11 As String
   Dim strTBD12 As String
   Dim strTBD13 As String
   Dim bolIsTit As Boolean
   Dim strTemp As String
   Dim intSpecRow As Integer
   Dim strTitle As String
   
On Error GoTo ERRORSECTION1
   
   WordEdit = True
   If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
   g_WordAp.Visible = True
   g_WordAp.WindowState = wdWindowStateMaximize 'wdWindowStateMinimize  wdWindowStateMaximize
   With g_WordAp
   
      If bolRetry = True Then
         m_intFileCnt = m_intFileCnt + 1
         'g_WordAp.Documents.Add.SaveAs m_WordFilePath & "\商標公報" & Left(txtTBD01, 2) & "卷" & Right(txtTBD01, 2) & "期" & "大陸公告明細" & Format(m_intFileCnt, "00") & ".doc"
         g_WordAp.Documents.add.SaveAs m_WordFilePath & "\商標公報" & Left(txtTBD01, 2) & "卷" & Right(txtTBD01, 2) & "期大陸公告明細.doc"
      End If
   
      If bolRetry = False Then .Selection.InsertBreak Type:=wdPageBreak
      
      .Selection.Font.Name = "標楷體"
      .Selection.PageSetup.Orientation = wdOrientPortrait
      .Selection.Orientation = wdTextOrientationHorizontal
      .Selection.Font.Size = 14
      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      'Modify by Morgan 2008/7/3
      '.Selection.PageSetup.TopMargin = .CentimetersToPoints(3.53)
      '.Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(4.1)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2.5)
      'end 2008/7/3
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphJustify
      .Selection.ParagraphFormat.DisableLineHeightGrid = True
      
            
'      'Add By Sindy Modify 2011/11/29
'      'Add by Morgan 2011/7/12 因為第 2 頁以後不要有信頭故改回放在本文
'      If PUB_ReadDB2File(stFileName, iPicNo) = True Then
'         Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'         oShape.ZOrder 4
'         oShape.LockAnchor = True
'         oShape.LockAspectRatio = -1
'         oShape.Width = .CentimetersToPoints(21)
'         oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'         oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'         oShape.Left = .CentimetersToPoints(0)
'         oShape.Top = .CentimetersToPoints(0.5)
'         If PUB_ReadDB2File(stFileName, iPicNo2) = True Then
'            Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=stFileName, LinkToFile:=False, SaveWithDocument:=True)
'            oShape.ZOrder 4
'            oShape.LockAnchor = True
'            oShape.LockAspectRatio = -1
'            oShape.Width = .CentimetersToPoints(21)
'            oShape.RelativeHorizontalPosition = wdRelativeHorizontalPositionPage
'            oShape.RelativeVerticalPosition = wdRelativeVerticalPositionPage
'            oShape.Left = .CentimetersToPoints(0)
'            oShape.Top = .CentimetersToPoints(27.3)
'         End If
'         .Selection.EndKey Unit:=wdStory
'      End If
            
      'Add by Morgan 2008/7/17 配合新的開窗定稿改固定行高
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly
      .Selection.ParagraphFormat.LineSpacing = 15
      'end 2008/7/17
            
'      .Selection.TypeParagraph 'Add by Morgan 2008/6/11 CFT 信頭比較高
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
      
      '靠左
      .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
'      If m_AppAddrZip = "" Then
'         .Selection.TypeParagraph
'      End If
'      .Selection.TypeText getAddrData
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "致：" & m_AppName
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'      .Selection.TypeText "敬啟者："
'      .Selection.TypeParagraph
'      .Selection.TypeParagraph
'
'      .Selection.TypeText "　　恭禧！台端／貴公司之商標已獲准註冊！將註冊公告三個月。依法台端／貴公司自註冊公告之日起取得商標權，專用期間10年。"
'      .Selection.TypeParagraph
      
      .Selection.Font.Size = 10
      
      m_TM01 = "": m_TM02 = "": m_TM03 = "": m_TM04 = "": custarea = "": custsales = ""
      
      strSql = "select tmbulletindata.*,tmbulletinowner.*,tm01,tm02,tm03,tm04 " & _
                 "From tmbulletindata,tmbulletinowner,Trademark,tmbulletin " & _
                "Where tbd02=tbor01 And tbd03=tbor06 And tbor02=1 " & _
                  "and tbd04=tm12(+) " & _
                  "and tbd02=tmbm01(+) and tbd03=tmbm02(+) " & _
                  "and tmbm06='" & strKey1 & "' and tbd15='B' and TBD16='1' " & _
               "order by tbd02,tbd03 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenDynamic
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         For i = 1 To rsTmp.RecordCount
            If m_TM01 = "" Then
               m_TM01 = Trim("" & rsTmp.Fields("TM01"))
               m_TM02 = Trim("" & rsTmp.Fields("TM02"))
               m_TM03 = Trim("" & rsTmp.Fields("TM03"))
               m_TM04 = Trim("" & rsTmp.Fields("TM04"))
            End If
            intSpecRow = 0: strTitle = ""
            strTBD01 = "" & rsTmp.Fields("TBD01")
            strTBD01_2 = Left(txtTBD01, 2) & "卷" & Format(Right(txtTBD01, 2), "00") & "期　" & ChangeWStringToTDateString(ChgTMBM07ToDate(strTBD01))
            strTBD02 = "" & rsTmp.Fields("TBD02")
            strTBD03 = "" & rsTmp.Fields("TBD03")
            If strTBD03 = "7" Or strTBD03 = "8" Then
               strTitle = "標章"
            Else
               strTitle = "商標"
            End If
            strTBD03_2 = GetTradeMarkName(strTBD03, 0)
            strTBD04 = "" & rsTmp.Fields("TBD04")
            strTBD05 = "" & rsTmp.Fields("TBD05")
            strTBD06 = "" & rsTmp.Fields("TBD06")
            strTBD07 = "" & rsTmp.Fields("TBD07")
            strTBD08 = "" & rsTmp.Fields("TBD08")
            strTBD09 = "" & rsTmp.Fields("TBD09")
            strTBD10 = "" & rsTmp.Fields("TBD10")
            strTBD11 = "" & rsTmp.Fields("TBD11")
            strTBD12 = "" & rsTmp.Fields("TBD12")
            strTBD13 = "" & rsTmp.Fields("TBD13")
            .Selection.TypeText "----------------------------------------------------------------------------------------------"
            .Selection.TypeParagraph
            .Selection.TypeText "註冊" & strTBD03_2 & "第" & strTBD02 & "號　申請案號：" & strTBD04 & "　" & strTBD01_2 & "　商標圖樣：" & strTBD05
            .Selection.TypeParagraph
            .Selection.TypeText "申請日期：" & strTBD06 '& "|#右代表圖#|"
            AddInPicToWordR g_WordAp, strTBD12 '插入圖檔
            .Selection.TypeParagraph
            If strTBD13 <> "" Then
               .Selection.TypeText "優先權日：" & strTBD13
               .Selection.TypeParagraph
               intSpecRow = intSpecRow + 1
            End If
            .Selection.TypeText strTitle & "名稱：" & strTBD07
            .Selection.TypeParagraph
            '商標權人資料
            strSql = "select * from tmbulletinowner Where tbor01='" & strTBD02 & "' and tbor06='" & strTBD03 & "' order by tbor02 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               For j = 1 To RsTemp.RecordCount
                  bolIsTit = False '尚無標題
                  If j > 1 Then intSpecRow = intSpecRow + 1
                  If "" & RsTemp.Fields("tbor03") <> "" Then
                     If bolIsTit = False Then
                        .Selection.TypeText strTitle & "權人：" & "" & RsTemp.Fields("tbor03")
                        bolIsTit = True '有標題了
                     Else
                        .Selection.TypeText "　　　　　" & "" & RsTemp.Fields("tbor03")
                     End If
                     .Selection.TypeParagraph
                  End If
                  If "" & RsTemp.Fields("tbor04") <> "" Then
                     If bolIsTit = False Then
                        .Selection.TypeText strTitle & "權人：" & "" & RsTemp.Fields("tbor04")
                        bolIsTit = True '有標題了
                        intSpecRow = intSpecRow + 1
                     Else
                        .Selection.TypeText "　　　　　" & "" & RsTemp.Fields("tbor04")
                     End If
                     .Selection.TypeParagraph
                  End If
                  If "" & RsTemp.Fields("tbor05") <> "" Then
                     If bolIsTit = False Then
                        .Selection.TypeText strTitle & "權人：" & "" & RsTemp.Fields("tbor05")
                        bolIsTit = True '有標題了
                     Else
                        .Selection.TypeText "　　　　　" & "" & RsTemp.Fields("tbor05")
                     End If
                     .Selection.TypeParagraph
                  End If
                  RsTemp.MoveNext
               Next j
            End If
            '商標權人資料 End
            If strTBD08 <> "" Then
               .Selection.TypeText "代理人：" & strTBD08
               .Selection.TypeParagraph
            End If
            .Selection.TypeText "權利期間：" & strTBD09
            .Selection.TypeParagraph
            .Selection.TypeText "審查人員：" & strTBD10
            .Selection.TypeParagraph
            If intSpecRow = 0 Then
               .Selection.TypeParagraph
               .Selection.TypeParagraph
            ElseIf intSpecRow = 1 Then
               .Selection.TypeParagraph
            End If
            If strTBD11 <> "" Then
               .Selection.TypeText strTBD11
               .Selection.TypeParagraph
               .Selection.TypeParagraph
            End If
            '商品資料
            strSql = "select * from tmbulletingoods Where tbg01='" & strTBD02 & "' and tbg07='" & strTBD03 & "' order by tbg02 asc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               For j = 1 To RsTemp.RecordCount
                  If "" & RsTemp.Fields("tbg03") <> "" Then
                     .Selection.TypeText "" & RsTemp.Fields("tbg03")
                     .Selection.TypeParagraph
                  End If
                  strTemp = Trim("" & RsTemp.Fields("tbg04")) & _
                            Trim("" & RsTemp.Fields("tbg05")) & _
                            Trim("" & RsTemp.Fields("tbg06")) & _
                            Trim("" & RsTemp.Fields("tbg08")) & _
                            Trim("" & RsTemp.Fields("tbg09")) & _
                            Trim("" & RsTemp.Fields("tbg10"))
                  If strTemp <> "" Then
                     If strTBD03 = "7" Then '證明標章
                        .Selection.TypeText "證明內容：" & strTemp
                     ElseIf strTBD03 = "8" Then '團體標章
                        .Selection.TypeText "表彰內容：" & strTemp
                     Else
                        .Selection.TypeText "商品或服務名稱：" & strTemp
                     End If
                     .Selection.TypeParagraph
                  End If
                  RsTemp.MoveNext
               Next j
            End If
            '商品資料 End
            
            '加註:已產生定稿
            strSql = "update TMBulletinData set TBD14='Y' where TBD02='" & strTBD02 & "' and TBD03='" & strTBD03 & "'"
            cnnConnection.Execute strSql
            
            
            '筆數為偶數時,接下一頁
            If i <> rsTmp.RecordCount And (i Mod 2) = 0 Then
               .Selection.InsertBreak Type:=wdPageBreak
            End If
            
            rsTmp.MoveNext
         Next i
'         .Selection.TypeText "----------------------------------------------------------------------------------------------"
'         .Selection.TypeParagraph
      End If
      
'      .Selection.Font.Size = 14
'
'      .Selection.TypeText "商標於註冊後應使用，否則連續三年無正當事由未使用，商標權將被廢止；為便於拓展外銷市場，宜於國內商標註冊後，申請大陸及其他各國商標之註冊。倘使　台端／貴公司對商標之使用，尚有質疑，敬祈不吝來電或蒞臨洽詢，本所二百多位專業人士竭誠為您提供最完善的服務！"
'      .Selection.TypeParagraph
'
'      .Selection.TypeParagraph
''      .Selection.TypeParagraph
'      .Selection.TypeText "　　　　耑此　　順頌"
'      .Selection.TypeParagraph
'      .Selection.TypeText "商　祺"
'      .Selection.TypeParagraph
'      .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　台一國際專利商標事務所　敬上"
'      .Selection.TypeParagraph
'      If m_TM01 <> "" Then
'         Call GetSales
'         If custarea = "" Then
'            .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　" & custsales
'         Else
'            .Selection.TypeText "　　　　　　　　　　　　　　　　　　　　　　　　" & custarea & "　" & custsales
'         End If
'         .Selection.TypeParagraph
'      End If
''      .Selection.TypeParagraph
''      .Selection.WholeStory
''      ChgWordFormat g_WordAp, .Selection.Text
   End With
   
'   PhaseIndent    '調整首行凸排
'   g_WordAp.Visible = True
'   g_WordAp.WindowState = wdWindowStateMaximize
   bolRetry = False
   Exit Function
   
ERRORSECTION1:
   If Err.Number <> 0 Then
      Select Case Err.Number
         Case 91, 462:
            Set g_WordAp = New Word.Application
            g_WordAp.Documents.add
            If bolRetry = False Then
               bolRetry = True
               Resume
            End If
         Case Else:
            MsgBox "錯誤 : " & Err.Description, vbCritical
            WordEdit = False
      End Select
   End If
End Function

Private Sub AddInPicToWordR(ByRef oWord As Word.Application, strFileName As String)
   Dim bytes() As Byte
   Dim file_num As Integer
   Dim rsPic As New ADODB.Recordset
   Dim IsWmf As Boolean
   Dim stSQL As String
   Dim intR As Integer
   Dim stFileName As String
   Dim oShape 'Added by Lydia 2016/09/29
   
On Error GoTo ErrHnd

   With oWord
      If InStr(strFileName, "imagesdata") = 0 Then
         strFileName = "imagesdata/" & strFileName
      End If

      '插入圖片檔案
      .ChangeFileOpenDirectory txtPath(0) & "\imagesdata\"
      '指定檔名
      'Modified by Lydia 2016/09/29 用舊寫法會造成Word2010出錯
      '.ActiveDocument.Shapes.AddPicture Anchor:=.Selection.Range, FileName:= _
      'txtPath(0) & "\" & strFileName, LinkToFile:= _
      'False, SaveWithDocument:=True
      '.ActiveDocument.Shapes("Picture " & Trim(.ActiveDocument.Shapes.Count + 1)).Select
      Set oShape = .ActiveDocument.Shapes.AddPicture(Anchor:=.Selection.Range, FileName:=txtPath(0) & "\" & strFileName, LinkToFile:=False, SaveWithDocument:=True)
      oShape.Select
      
      '定義大小
      '鎖定最高 圖區
      '圖大小
      'Modified by Lydia 2016/09/29
'      .Selection.ShapeRange.LockAspectRatio = msoTrue
'      .Selection.ShapeRange.Height = 230
'      If .Selection.ShapeRange.Width > 150 Then
'         .Selection.ShapeRange.Width = 150
'      End If
'      '移到指定位置
'      .Selection.ShapeRange.Left = .CentimetersToPoints(12) '11.2
'      '.Selection.ShapeRange.Top = .CentimetersToPoints(1)
'      .Selection.ShapeRange.LockAnchor = False
'      '圖蓋文
'      .Selection.ShapeRange.WrapFormat.Type = wdWrapSquare 'wdWrapNone.圖蓋文 wdWrapSquare.文字繞圖
      oShape.LockAspectRatio = msoTrue
      oShape.Height = 230
      If oShape.Width > 150 Then
         oShape.Width = 150
      End If
      '移到指定位置
      oShape.Left = .CentimetersToPoints(12)
      oShape.LockAnchor = False
      '圖蓋文
      oShape.WrapFormat.Type = wdWrapSquare
      
      .Selection.EndKey Unit:=wdStory
   End With
   Exit Sub
   
'加判斷若錯誤為無法刪除檔案時繼續(下次跑整批定稿時會刪)
ErrHnd:
   If (pub_OS = 1 And Err.Number = 75) Or (pub_OS <> 1 And Err.Number = 70) Then 'Err.Number = 5152
      Resume Next
   Else
      Err.Raise Err.Number
   End If
End Sub

'調整首行凸排
Sub PhaseIndent()
    g_WordAp.Selection.WholeStory
    With g_WordAp.Selection.ParagraphFormat
        .LeftIndent = g_WordAp.CentimetersToPoints(1)
        .RightIndent = g_WordAp.CentimetersToPoints(0)
        .SpaceBefore = 0
        .SpaceAfter = 0
        .LineSpacingRule = wdLineSpaceExactly
        .LineSpacing = 15
        .Alignment = wdAlignParagraphJustify
        .WidowControl = False
        .KeepWithNext = False
        .KeepTogether = False
        .PageBreakBefore = False
        .NoLineNumber = False
        .Hyphenation = True
        .FirstLineIndent = g_WordAp.CentimetersToPoints(-1)
        .OutlineLevel = wdOutlineLevelBodyText
        .AutoAdjustRightIndent = True
        .DisableLineHeightGrid = True
        .FarEastLineBreakControl = True
        .WordWrap = True
        .HangingPunctuation = True
        .HalfWidthPunctuationOnTopOfLine = False
        .AddSpaceBetweenFarEastAndAlpha = True
        .AddSpaceBetweenFarEastAndDigit = True
        .BaseLineAlignment = wdBaselineAlignAuto
    End With
End Sub

'將公報卷期轉換為日期
Private Function ChgTMBM07ToDate(strData As String)
Dim strYY As String
Dim strMM As String
Dim strDD As String
'920101 : 3001, 920116 : 3002 ...(每年會有24期)

strYY = (Val(Mid(strData, 1, Len(strData) - 2)) + 62)
strMM = Format(Right(strData, 2) / 2, "00")
If Right(strData, 2) Mod 2 <> 0 Then
    strDD = "01"
Else
    strDD = "16"
End If
ChgTMBM07ToDate = DBDATE(strYY & strMM & strDD)
End Function

Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   QueryData = False
   Label7 = "( 0 筆)"
   txtTBD01 = ""
   
   strSql = "select count(*) from TMBulletinData " & _
            "Where tbd15='B' and TBD16='1' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields(0)) Then
         If Val(rsTmp.Fields(0)) > 0 Then
            QueryData = True
            Label7 = "( " & rsTmp.Fields(0) & " 筆)"
            
            rsTmp.Close
            strSql = "SELECT distinct tbd01 FROM TMBulletinData WHERE tbd15='B' and TBD16='1' "
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               txtTBD01 = "" & rsTmp.Fields(0)
            End If
         End If
      End If
   End If
   Set rsTmp = Nothing
   
   If QueryData = False Then
      MsgBox "無資料！", vbOKOnly, "商標公報大陸清單列印"
   End If
End Function
