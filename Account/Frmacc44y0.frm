VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frmacc44y0 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "補扣繳地址條及清單"
   ClientHeight    =   5330
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5330
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7530
      TabIndex        =   5
      Top             =   240
      Width           =   765
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "印地址條"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6360
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5490
      TabIndex        =   3
      Top             =   240
      Width           =   765
   End
   Begin VB.ComboBox Combo3 
      Height          =   300
      Left            =   1740
      Style           =   2  '單純下拉式
      TabIndex        =   11
      Top             =   4980
      Width           =   3450
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3975
      Left            =   30
      TabIndex        =   2
      Top             =   930
      Width           =   8865
      _ExtentX        =   15646
      _ExtentY        =   7003
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   "不印|收據抬頭|繳款書寄件處|地址|收件人"
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
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   330
      Left            =   2190
      TabIndex        =   0
      Top             =   480
      Width           =   1275
      _ExtentX        =   2258
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   330
      Left            =   3630
      TabIndex        =   1
      Top             =   480
      Width           =   1275
      _ExtentX        =   2258
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox3 
      Height          =   330
      Left            =   2190
      TabIndex        =   6
      Top             =   90
      Width           =   1275
      _ExtentX        =   2258
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox4 
      Height          =   330
      Left            =   3630
      TabIndex        =   7
      Top             =   90
      Width           =   1275
      _ExtentX        =   2258
      _ExtentY        =   582
      _Version        =   393216
      BackColor       =   -2147483633
      AllowPrompt     =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   5760
      TabIndex        =   12
      Top             =   4950
      Width           =   2745
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      BackStyle       =   0  '透明
      Caption         =   "地址條印表機："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5040
      Width           =   1485
   End
   Begin VB.Line Line1 
      X1              =   3360
      X2              =   3570
      Y1              =   630
      Y2              =   630
   End
   Begin VB.Line Line2 
      X1              =   3360
      X2              =   3570
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "上次補扣繳日期"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   9
      Top             =   150
      Width           =   1545
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "補扣繳日期："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   810
      TabIndex        =   8
      Top             =   540
      Width           =   1335
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc44y0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/03/01 改成Form2.0 ; 地址條改成Excel列印
'Memo By Sindy 2022/2/17 Form2.0已修改 (Printer列印未改)
'Memo by Lydia 2022/02/11 改成Form2.0 ; GRD1改字型=新細明體-ExtB ; Printer列印未改
'Create By Sindy 2016/11/18
Option Explicit

Dim strPrinter As String
Dim iRow As Integer, iCol As Integer
Dim m_dftColor As Long '預設顏色
Dim m_dftColor2 As Long '預設顏色2
Dim m_dftColor3 As Long '點選顏色

Private Function TxtValidate() As Boolean
   TxtValidate = False

   '日期檢查
   If DateCheck(MaskEdBox1.Text) = MsgText(603) Then
      MsgBox "補扣繳起始日期格式錯誤！", vbExclamation
      TxtValidate = False
      MaskEdBox1.SetFocus
      Exit Function
   End If

   If DateCheck(MaskEdBox2.Text) = MsgText(603) Then
      MsgBox "補扣繳迄止日期格式錯誤！", vbExclamation
      TxtValidate = False
      MaskEdBox2.SetFocus
      Exit Function
   End If

   TxtValidate = True
End Function

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   '                        0       1           2               3       4
   arrGridHeadText = Array("要印", "收據抬頭", "繳款書寄件處", "地址", "收件人")
   arrGridHeadWidth = Array(450, 2200, 1100, 2500, 2200)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
Dim bolhavdData As Boolean
Dim ii As Integer
   
   If GRD1.Rows - 1 > 0 Then
      If GRD1.TextMatrix(1, 1) = "" Then
         MsgBox "無資料!!", vbCritical
         Exit Sub
      Else
         bolhavdData = False
         For ii = 1 To GRD1.Rows - 1
            If Trim(GRD1.TextMatrix(ii, 0)) = "V" Then '要印
               bolhavdData = True
               Exit For
            End If
         Next ii
      End If
      If bolhavdData = False Then
         MsgBox "無列印地址條的資料!!", vbCritical
         Exit Sub
      End If
      
      If MsgBox("是否要列印地址條？" & vbCrLf & _
                "若要印，請放地址條貼紙於選取的印表機!!", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         PUB_SetOsDefaultPrinter Combo3 'Added by Lydia 2022/03/01 切換Word/Excel印表機
         PUB_RestorePrinter Combo3
         PrintAddress '列印地址條
         PUB_SetOsDefaultPrinter strPrinter 'Added by Lydia 2022/03/01 切換Word/Excel印表機
         PUB_RestorePrinter strPrinter
         MsgBox "列印完畢!!", vbInformation
      End If
      
      PUB_SaveLastDate Me.Name, "MaskEdBox3", ChangeTDateStringToTString(MaskEdBox1)
      PUB_SaveLastDate Me.Name, "MaskEdBox4", ChangeTDateStringToTString(MaskEdBox2)
      MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox3"))
      MaskEdBox4.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
      MaskEdBox1.Mask = ""
      MaskEdBox2.Mask = ""
      MaskEdBox1.Text = ""
      MaskEdBox2.Text = ""
      MaskEdBox1.Mask = DFormat
      MaskEdBox2.Mask = DFormat
   End If
End Sub

Private Sub PrintAddress()
   Dim ii As Integer
   Dim strAddr As String, strCustName As String
   Dim strTempAddressList As String 'Added by Lydia 2022/03/01
   
   Screen.MousePointer = vbHourglass
   For ii = 1 To GRD1.Rows - 1
      If Trim(GRD1.TextMatrix(ii, 0)) = "V" Then '要印
         strAddr = GRD1.TextMatrix(ii, 3)
         strCustName = GRD1.TextMatrix(ii, 4)
         'Modified by Lydia 2022/03/01 傳入多張地址條的內容；用|區隔不同張地址條，同一張地址條用$區隔地址和收件人
         'Call PUB_PrintAccAddress(strAddr, strCustName)
         If strAddr & strCustName <> "" Then strTempAddressList = strTempAddressList & Trim(strAddr) & "$" & Trim(strCustName) & "|"
      End If
   Next ii
   'Added by Lydia 2022/03/01 改用Execl列印地址條
   If strTempAddressList <> "" Then
       If PUB_XlsAccAddress(strTempAddressList) = False Then
           MsgBox "列印失敗！", vbCritical
       End If
   End If
   'end 2022/03/01
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdQuery_Click()
   If TxtValidate Then
      Screen.MousePointer = vbHourglass
      Call doQuery
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Grd1_Click()
   Dim iCurCol As Integer, iCurRow As Integer
   
   With GRD1
   If .MouseRow > 0 And .MouseRow < .Rows And .MouseCol < .Cols Then
      iCurRow = .MouseRow
      iCurCol = .MouseCol
      .Visible = False
      
      .row = iCurRow
      .col = 0
      If Trim(.TextMatrix(.row, .col)) = "" Then
         .TextMatrix(.row, .col) = "V"
         SetColor iCurRow, m_dftColor3
      Else
         .TextMatrix(.row, .col) = ""
         SetColor iCurRow, m_dftColor
      End If
      
      .col = iCurCol
      iRow = .row: iCol = .col
           
      .Visible = True
   End If
   End With
End Sub

Private Sub SetColor(pRow As Integer, pColor As Long)
   With GRD1
   .row = pRow
   For intI = 0 To .Cols - 1
      .col = intI
      .CellBackColor = pColor
   Next
   End With
End Sub

Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
GRD1.ToolTipText = ""
If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
   If GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol) <> "" Then
      GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
   End If
End If
End Sub

Private Sub doQuery()
Dim rs As ADODB.Recordset
Dim strT15 As String, ii As Integer
Dim m_CU01 As String, m_CU02 As String
Dim m_CU168 As String, m_CU169 As String
Dim m_CU170 As String, m_CU171 As String
   
   Screen.MousePointer = vbHourglass
   Label4 = ""
   
   '產生暫存檔資料
   '**********************************************
   'Acctmp08:
   '**********************************************
   'T01:流水號 key
   'T02:'' key
   'T05:Me.Name key
   'T06:a1p18入帳日期 key
   'T14:strUserName key
   'T15:a1p04收據抬頭
   '**********************************************
   adoTaie.Execute "delete from ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   '讀取基礎資料
   'Modify By Sindy 2020/5/12 and substr(a1p05,1,4)='1102' => and instr('1102,1105,1106',substr(a1p05,1,4))>0
   'Modify By Sindy 2020/6/5 and a1p01='1' => and a1p01<>'J'
   adoTaie.Execute "insert into acctmp08(T01,T02,T05,T06,T14,T15)" & _
                   " SELECT distinct rownum,' ','" & Me.Name & "',a1p18,'" & strUserNum & "',a1p04" & _
                   " From Acc1p0" & _
                   " Where A1P18 >= " & ACDate(DBDATE(MaskEdBox1)) & " And A1P18 <= " & ACDate(DBDATE(MaskEdBox2)) & _
                   " and a1p01<>'J'" & _
                   " and a1p02='E'" & _
                   " and a1p08>0" & _
                   " and instr('1102,1105,1106',substr(a1p05,1,4))>0 and substr(a1p04,1,1)<>'K'"
   '解析收據抬頭
   strExc(0) = "select T01,T06,T15" & _
               " From ACCTMP08" & _
               " where T05='" & Me.Name & "' and T14='" & strUserNum & "'"
   intI = 1
   Set rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rs.MoveFirst
      Do While Not rs.EOF
         strT15 = rs.Fields("T15")
         'modify by sonia 2023/1/6 不能用INSTR，因為112/1/5補扣繳111年之第2筆會是"鼎建實業股份有限公司1112",截取結果會變成"鼎建實業股份有限公司1"造成抓不到地址
         'If InStr(rs.Fields("T15"), Left(rs.Fields("T06"), 3)) > 0 Then
         If Left(Right(rs.Fields("T15"), 4), 3) = Left(rs.Fields("T06"), 3) Then
            strT15 = Mid(rs.Fields("T15"), 1, InStr(rs.Fields("T15"), Left(rs.Fields("T06"), 3)) - 1)
         Else
            For ii = 1 To Len(rs.Fields("T15"))
               If IsNumeric(Mid(rs.Fields("T15"), ii, 1)) = True Then
                  strT15 = Mid(rs.Fields("T15"), 1, ii - 1)
                  Exit For
               End If
            Next ii
         End If
         
         If GetTitleCustData(strT15, "", "", m_CU01, m_CU02, _
                            "", "", "", "", "", "", "", _
                            "", "", "", "", "", "", "", _
                            "", "", "", , m_CU168, m_CU169, m_CU170, m_CU171, "") = True Then
         End If
         'Modify By Sindy 2017/10/16 瑞婷說不排除"有"每月提醒代填繳款書,ex:鐵雲
'         If m_CU168 = "Y" Then '每月提醒代填繳款書,則刪除
'            adoTaie.Execute "delete from ACCTMP08 where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T15='" & rs.Fields("T15") & "'"
'         Else
            '更新資料:
            'T20:繳款書地址 m_CU170
            'T16:收件人     m_CU171
            'T25:繳款書寄件處 m_CU169
            'T02.客戶編號
            strSql = "update ACCTMP08 set" & _
                     " T15=" & CNULL(strT15) & _
                     ",T20=" & CNULL(m_CU170) & _
                     ",T16=" & CNULL(m_CU171) & _
                     ",T25=" & CNULL(m_CU169) & _
                     ",T02=" & IIf(m_CU01 & m_CU02 = "", "' '", CNULL(m_CU01 & m_CU02)) & _
                     " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
                     " and T15='" & rs.Fields("T15") & "'"
            cnnConnection.Execute strSql, intI
'         End If
         rs.MoveNext
      Loop
   End If
   
   GRD1.Clear
   SetGrd
   strExc(0) = "select distinct ' ' 要印,T15 收據抬頭,decode(T25,'T','收據抬頭','C','客戶','2','會計師','3','特殊','') 繳款書寄件處,T20 地址,T16 收件人" & _
               " from ACCTMP08" & _
               " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " order by T15"
   intI = 1
   Set rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      MsgBox "查無資料！", , MsgText(5)
   Else
      Set GRD1.Recordset = rs
      Label4 = "查詢出 " & rs.RecordCount & " 筆"
   End If
   
   Set rs = Nothing
   Screen.MousePointer = vbDefault
End Sub

Private Sub ClearAll()
   MaskEdBox1.Mask = ""
   MaskEdBox2.Mask = ""
   MaskEdBox1.Text = ""
   MaskEdBox2.Text = ""
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   
   '上次發放日期
   If PUB_GetLastDate(Me.Name, "MaskEdBox3") <> "" Then
      MaskEdBox3.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox3"))
   End If
   If PUB_GetLastDate(Me.Name, "MaskEdBox4") <> "" Then
      MaskEdBox4.Text = ChangeTStringToTDateString(PUB_GetLastDate(Me.Name, "MaskEdBox4"))
   End If
   MaskEdBox3.Mask = DFormat
   MaskEdBox4.Mask = DFormat
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single

   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 9045
   Me.Height = 5700
   '改單線固定(調整大小不用再設定)
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
      For intY = 0 To Int(ScaleHeight / sglHeight)
         PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
      Next
   Next

   Call ClearAll
   Call SetGrd
   PUB_SetPrinter Me.Name, Combo3, strPrinter
   
   '底色
   m_dftColor = &HFFFFFF
   '底色2
   m_dftColor2 = RGB(&HFF, &HFA, &HCD)
   '底色3
   m_dftColor3 = &HFFC0C0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '若印表機變動, 則更新列印設定
   If Me.Combo3.Text <> Me.Combo3.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo3.Name, "0", "0", Me.Combo3.Text
   End If

   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc44y0 = Nothing
End Sub
