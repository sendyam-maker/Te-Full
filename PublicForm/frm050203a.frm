VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050203a 
   BorderStyle     =   1  '單線固定
   Caption         =   "未請款明細查詢"
   ClientHeight    =   5710
   ClientLeft      =   2520
   ClientTop       =   2570
   ClientWidth     =   9320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5710
   ScaleWidth      =   9320
   Begin VB.CommandButton Command1 
      Caption         =   "列印(P)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   6750
      TabIndex        =   3
      Top             =   15
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MFG1 
      Height          =   5208
      Left            =   36
      TabIndex        =   1
      Top             =   468
      Width           =   9252
      _ExtentX        =   16334
      _ExtentY        =   9172
      _Version        =   393216
      Cols            =   16
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
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
      _Band(0).Cols   =   16
   End
   Begin VB.CommandButton Command1 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   0
      Left            =   8040
      TabIndex        =   0
      Top             =   20
      Width           =   1200
   End
   Begin VB.Label lbl 
      Height          =   180
      Index           =   1
      Left            =   3000
      TabIndex        =   4
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label lbl 
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "frm050203a"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/09/22 改成Form2.0 ; MFG1改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/2 日期欄已修改
Option Explicit

Dim strSql As String, i As Integer, j As Integer, s As Integer, strTemp As Variant
Dim StrTest As String, StrTest1 As String, intK As Integer, strSQL1 As String, strSQL2 As String, StrSQL3 As String, StrSQL4 As String, strSQL5 As String, StrSQL6 As String
Dim Page As Integer, iPrint As Integer
Dim strTemp3, PLeft 'Modify by Amy 2016/12/07 原:strTemp3(0 To 9) As String,PLeft(0 To 9) As Integer
'Add by Amy 2013/11/20
Dim strTp(2) As String
Dim FMPFFPstate As String
Dim strPA150 As String
'Add by Amy 2016/12/07
Dim arrField
Dim bolFCPFG As Boolean  '是否有查FCP/FG資料(新格式)
Dim iMaxHeight As Integer   'Added by Lydia 2018/02/22

Private Sub Command1_Click(Index As Integer)
  Select Case Index
  Case 0
        Me.Hide
  Case 1
       Screen.MousePointer = vbHourglass
       'Add by Amy 2016/12/07 有查FCP/FG 報表格式不同
       bolFCPFG = False: strExc(0) = "智權人員"
       If InStr(frm050203.Text1(0), "FCP") > 0 Or InStr(frm050203.Text1(0), "FG") > 0 Then
            bolFCPFG = True
            ReDim arrField(10)
            ReDim PLeft(10)
            ReDim strTemp3(10)
            If InStr(1, frm050203.Text1(0), "FCP") > 0 Then strExc(0) = "承辦業務"
            arrField = Array("承辦人", strExc(0), "本所案號", "案件性質", "申請國家", "FC代/申請人國籍", "未請款逾月數", "發文日", "FC代理人", "發文規費", "帳單金額")
            PLeft = Array(300, 1500, 2600, 4800, 6600, 7900, 9200, 10200, 11700, 14300, 15700)
       Else
            ReDim arrField(11)
            ReDim PLeft(11)
            ReDim strTemp3(11)
            arrField = Array("承辦人", strExc(0), "本所案號", "案件名稱", "案件性質", "申請國家", "FC代/申請人國籍", "未請款逾月數", "發文日", "FC代理人", "發文規費", "帳單金額")
            PLeft = Array(300, 1400, 2400, 4300, 6000, 7700, 8900, 10100, 11100, 12300, 14700, 15800)
       End If
       'end 2016/12/14
       '910802 列印功能   nick 新撰寫
       PrintData
       Screen.MousePointer = vbDefault
  End Select
End Sub

Sub PrintData()
   'Add by Amy 2016/12/07
   Dim j As Integer, intShowCol As Integer, strTp As String
   Dim iOrt As Integer  'Added by Lydia 2018/02/22
   
   If MFG1.Rows <> 1 Then
       Page = 1
       'Add by Amy 2013/11/20
       strPA150 = ""
       If InStr(1, frm050203.Text1(0), "FCP") > 0 Then
            strPA150 = MFG1.TextMatrix(1, 11)
       End If
       'end 2013/11/20
       'Added by Lydia 2018/02/22 設定紙張和方向
       iOrt = Printer.Orientation
       Printer.PaperSize = 9 'A4
       Printer.Orientation = 2
       iMaxHeight = Printer.ScaleHeight - 1500
       'end 2018/02/22
       PrintTitle
       Dim intItem As Long  'Modify by Amy 2013/11/20 原:Integer (查 P,PS 2帳單 920101-1020731 有206905筆)
       With MFG1
           For intItem = 1 To .Rows - 1
               'Modify by Amy 2013/11/20 +if
               If InStr(1, frm050203.Text1(0), "FCP") > 0 Then
                    If strPA150 <> .TextMatrix(intItem, 11) Then
                        SetPSWord 'Add by Amy 2016/12/07
                        strPA150 = .TextMatrix(intItem, 11)
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
               End If
               'end 2013/11/20
               'Modified by Lydia 2018/02/22
               'If iPrint >= 10000 Then
               If iPrint >= iMaxHeight Then
                   SetPSWord 'Add by Amy 2016/12/07
                   Printer.NewPage
                   Page = Page + 1
                   PrintTitle
               End If
               .row = intItem
               'Modify by Amy 2016/12/07
'               .col = 0
'               '2012/5/1 modify by sonia 字太長
'               'strTemp3(0) = .Text
'               strTemp3(0) = StrToStr(.Text, 4)
'               .col = 1
'               strTemp3(1) = .Text
'               .col = 2
'               strTemp3(2) = .Text
'               .col = 3
'               strTemp3(3) = StrToStr(.Text, 9)
'               .col = 4
'               '911113 nick 修正太長
'               'strTemp3(4) = .Text
'               strTemp3(4) = StrToStr(.Text, 4)
'               .col = 5
'               strTemp3(5) = .Text
'               .col = 6
'               strTemp3(6) = StrToStr(.Text, 4) 'Modify by Amy 2016/12/07
'               .col = 7
'               strTemp3(7) = .Text
'               .col = 8
'               strTemp3(8) = .Text
'               .col = 9
'               strTemp3(9) = StrToStr(.Text, 9)
               For j = 0 To UBound(arrField)
'                    If .TextMatrix(0, j) <> "" Then
                        Select Case j
                            Case GetValue("承辦人")
                                strTp = StrToStr(.TextMatrix(intItem, GetValue("承辦人", True)), 4)
                            Case GetValue("案件性質")
                                strTp = StrToStr(.TextMatrix(intItem, GetValue("案件性質", True)), 6)
                            Case GetValue("FC代/申請人國籍")
                                strTp = StrToStr(.TextMatrix(intItem, GetValue("FC代/申請人國籍", True)), 4)
                            Case GetValue("FC代理人")
                                strTp = StrToStr(.TextMatrix(intItem, GetValue("FC代理人", True)), 9)
                            Case Else
                                 strExc(0) = arrField(j)
                                If bolFCPFG = False Then
                                    strTp = .TextMatrix(intItem, GetValue(strExc(0), True))
                                    If j = GetValue("案件名稱") Then strTp = StrToStr(strTp, 6)
                                Else
                                    strTp = .TextMatrix(intItem, GetValue(strExc(0), True))
                                End If
                        End Select
                        strTemp3(j) = strTp
'                        If bolFCPFG = True And j = GetValue("案件名稱") Then
'                            '有查 FCP或FG 報表不印案件名稱
'                        Else
'                            strTemp3(intShowCol) = strTp
'                            intShowCol = intShowCol + 1
'                        End If
'                    End If
               Next j
'               intShowCol = 0
               'end 2016/12/14
               PrintDatil
           Next intItem
       End With
       SetPSWord 'Add by Amy 2016/12/07
       Printer.EndDoc
       ShowPrintOk
       Printer.Orientation = iOrt 'Added by Lydia 2018/02/22
   Else
       MsgBox "沒有資料可以列印 !", vbCritical
   End If
End Sub

Sub PrintTitle()
   'GetPleft 'Mark by Amy 2016/12/07 不使用
   iPrint = 500
   'Printer.Orientation = 2 'Remove by Lydia 2018/02/22
   Printer.Font.Name = "細明體"
   'Modified by Lydia 2018/02/22
   'Printer.Font.Size = 22
   Printer.Font.Size = 20
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 7500 - (Printer.TextWidth("未請款明細查詢") / 2)
   Printer.CurrentY = iPrint
   Printer.Print "未請款明細查詢"
   iPrint = iPrint + 500
   'Modified by Lydia 2018/02/22
   'Printer.Font.Size = 12
   Printer.Font.Size = 11
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 7500 - (Printer.TextWidth("系統類別：" & frm050203.Text1(0).Text & strTp(1)) / 2)
   Printer.CurrentY = iPrint
   Printer.Print "系統類別：" & frm050203.Text1(0).Text & strTp(1)
   iPrint = iPrint + 300
   Printer.CurrentX = 7500 - (Printer.TextWidth("發文日：" & Format(ChangeTStringToTDateString(frm050203.Text1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm050203.Text1(2))) / 2)
   Printer.CurrentY = iPrint
   Printer.Print "發文日：" & Format(ChangeTStringToTDateString(frm050203.Text1(1)) & " ", "@@@@@@@@@") & "－" & ChangeTStringToTDateString(frm050203.Text1(2))
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   'Add by Amy 2013/11/20 +if
   If InStr(1, frm050203.Text1(0), "FCP") > 0 And strPA150 <> MsgText(601) Then
        '工程師組別
        Printer.Print "工程師組別：" & PUB_GetFCPGrpName(strPA150)
   Else
        Printer.Print lbl(0).Caption
   End If
   '93.6.16 ADD BY SONIA
   Printer.CurrentX = 6500
   Printer.CurrentY = iPrint
   Printer.Print lbl(1).Caption
   '93.6.16 END
   Printer.CurrentX = 13000
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(Page)
   iPrint = iPrint + 300
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   'Modify by Amy 2016/12/07 查FCP/FG 欄位顯示不同
'   Printer.CurrentX = PLeft(0)
'   Printer.CurrentY = iPrint
'   Printer.Print "承辦人"
'   Printer.CurrentX = PLeft(1)
'   Printer.CurrentY = iPrint
'   'Add by Amy 2013/11/20 +if
'   If InStr(1, frm050203.Text1(0), "FCP") > 0 Then
'        Printer.Print "承辦業務"
'   Else
'        Printer.Print "智權人員"
'   End If
'   Printer.CurrentX = PLeft(2)
'   Printer.CurrentY = iPrint
'   Printer.Print "本所案號"
'   Printer.CurrentX = PLeft(3)
'   Printer.CurrentY = iPrint
'   Printer.Print "案件名稱"
'   Printer.CurrentX = PLeft(4)
'   Printer.CurrentY = iPrint
'   Printer.Print "案件性質"
'   Printer.CurrentX = PLeft(5)
'   Printer.CurrentY = iPrint
'   Printer.Print "申請國家"
'   Printer.CurrentX = PLeft(6)
'   Printer.CurrentY = iPrint
'   Printer.Print "FC代/申請人國籍"  'Modify by Amy 2013/11/20 原:申請人/代理人國籍
'   Printer.CurrentX = PLeft(7)
'   Printer.CurrentY = iPrint
'   Printer.Print "收文日"
'   Printer.CurrentX = PLeft(8)
'   Printer.CurrentY = iPrint
'   Printer.Print "發文日"
'   Printer.CurrentX = PLeft(9)
'   Printer.CurrentY = iPrint
'   Printer.Print "FC代理人"
'    iPrint = iPrint + 300
   For i = LBound(PLeft) To UBound(PLeft)
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        If i = GetValue("FC代/申請人國籍") Then
            Printer.Print "FC 代/申"
            Printer.CurrentX = PLeft(i)
            Printer.CurrentY = iPrint + 300
            Printer.Print "請人國籍"
        ElseIf i = GetValue("未請款逾月數") Then
            Printer.Print "未請款"
            Printer.CurrentX = PLeft(i)
            Printer.CurrentY = iPrint + 300
            Printer.Print "逾月數"
        Else
            Printer.Print arrField(i)
        End If
   Next i
   iPrint = iPrint + 600
   'end 2016/12/14
   
   Printer.CurrentX = 300
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
End Sub

'Mark by Amy 2016/12/07 不使用
'Sub GetPleft()
'   Erase PLeft
'   PLeft(0) = 500
'   PLeft(1) = 1500
'   PLeft(2) = 2500
'   PLeft(3) = 4700
'   PLeft(4) = 7000
'   PLeft(5) = 8200
'   PLeft(6) = 9400
'   PLeft(7) = 11700
'   PLeft(8) = 13000
'   PLeft(9) = 14200
'End Sub

Sub PrintDatil()
   For i = LBound(PLeft) To UBound(PLeft)
       If i = GetValue("未請款逾月數") Then
            Printer.CurrentX = PLeft(i) + (400 - Printer.TextWidth(strTemp3(i)))
       ElseIf i = GetValue("發文規費") Then
            Printer.CurrentX = PLeft(i + 1) - 300 - Printer.TextWidth(strTemp3(i))
       ElseIf i = GetValue("帳單金額") Then
            Printer.CurrentX = PLeft(i) + (900 - Printer.TextWidth(strTemp3(i)))
       Else
            Printer.CurrentX = PLeft(i)
       End If
       Printer.CurrentY = iPrint
       Printer.Print strTemp3(i)
   Next i
   iPrint = iPrint + 300
End Sub


Private Sub Form_Load()
'edit by nickc 2007/02/06 不用 dll 了 Dim obj01 As Object

Dim i As Integer
    MoveFormToCenter Me
    'MFG1.CellAlignment = 9
    MFG1.Rows = 2
    'MFG1.Cols = 9
    MFG1.FixedRows = 1
    MFG1.FixedCols = 0
    MFG1.ColWidth(0) = 1000
    MFG1.ColWidth(1) = 1000
    MFG1.ColWidth(2) = 1550
    MFG1.ColWidth(3) = 2000
    MFG1.ColWidth(4) = 1000
    MFG1.ColWidth(5) = 1000
    MFG1.ColWidth(6) = 1600
    MFG1.ColWidth(7) = 800
    MFG1.ColWidth(8) = 800
    MFG1.ColWidth(9) = 1000
    MFG1.ColWidth(10) = 0
    MFG1.ColWidth(11) = 0
    MFG1.ColWidth(12) = 0
    MFG1.ColWidth(13) = 0
    'Modify by Amy 2016/12/07
    MFG1.ColWidth(14) = 1000
    MFG1.ColWidth(15) = 1000
    'end 2016/12/07
    'StrSQL = frm050203.Text2
    
    'Set Rss = objPublicData.ReadRst(StrSQL, True)
    'Set MFG1.Recordset = Rss
    With MFG1
    
        .TextMatrix(0, 0) = "承辦人"
        .TextMatrix(0, 1) = "智權人員"
        .TextMatrix(0, 2) = "本所案號"
        .TextMatrix(0, 3) = "案件名稱"
        .TextMatrix(0, 4) = "案件性質"
        .TextMatrix(0, 5) = "申請國家"
        .TextMatrix(0, 6) = "申請人/代理人國籍"
        .TextMatrix(0, 7) = "收文日"
        .TextMatrix(0, 8) = "發文日"
        .TextMatrix(0, 9) = "FC代理人"
        .TextMatrix(0, 10) = ""
        .TextMatrix(0, 11) = ""
        .TextMatrix(0, 12) = ""
        .TextMatrix(0, 13) = ""
        'Modify by Amy 2016/12/07
        .TextMatrix(0, 14) = "發文規費"
        .TextMatrix(0, 15) = "帳單金額"
        'end 2016/12/07
   End With
   'Add By Cheng 2002/04/24
   If frm050203.Text1(9).Text = "1" Then
      Me.lbl(0).Caption = "查詢順序：智權人員"
   Else
      Me.lbl(0).Caption = "查詢順序：發文日"
   End If
   '93.6.16 ADD BY SONIA
   If frm050203.Text1(11).Text = "1" Then
      Me.lbl(1).Caption = "單據類別：收據 / 請款單"
   Else
      Me.lbl(1).Caption = "單據類別：帳單"
   End If
End Sub
    
Sub StrMenu()
'Add By Cheng 2002/07/09
Dim StrSQLa As String
Dim StrSqlB As String
'Add by Amy 2013/11/20
Dim strFMPFFP1 As String, strFMPFFP2 As String, strFMPFFP3 As String
Dim lngQ As Long 'Added by Lydia 2018/02/22

   'Screen.MousePointer = vbHourglass
   Me.Enabled = False
   strSQL1 = ""
   strSQL2 = ""
   StrSQL3 = ""
   StrSQL4 = ""
   strSQL5 = ""
   'Add by Amy 2013/11/20 判斷是否含/排除 FMP、FFP
   strFMPFFP1 = "": strFMPFFP2 = "": strFMPFFP3 = "": FMPFFPstate = "": strTp(1) = ""
   If Len(Trim(Me.Tag)) > 0 Then
        FMPFFPstate = Left(Me.Tag, 3)
        If FMPFFPstate = "ADD" Then
            '含 FMP、FFP
            strFMPFFP1 = " And CP01 In ('P','CFP','') And Substr(CP12,1,1)='F' And CP10<>'404' "
            strFMPFFP2 = " And CP01 In ('PS','CPS','') And Substr(CP12,1,1)='F' And CP10<>'404' "
            strTp(1) = "　(含 FMP、FFP)"
        ElseIf FMPFFPstate = "DEL" Then
            strTp(1) = "　(不含 FMP、FFP)"
        End If
   End If
   'end 2013/11/20
   
    'Add by Lydia 2014/11/14 FCP承辦區域特殊狀況之智權人員劃分方式
    '代理人Y51333010=Pub_GetSpecMan("北京銀龍FCP案承辦業務") ,NA51 = decode(pa75," & midstr & ",na51)
     Dim midStr As String
     'Modified by Lydia 2016/02/03改成回傳case句
     'midStr = Pub_GetSpecMan("北京銀龍FCP案承辦業務")
     midStr = Pub_GetSpecFCP
   
   'Add By Cheng 2003/09/23
   'FCP,FG不印案件性質不續辦(907), 閉卷(913), 延期(404)
   'Begin
   strSQL1 = strSQL1 + " AND CP01||CP10 Not In ('FCP907', 'FCP913', 'FCP404', 'FG907', 'FG913', 'FG404') "
   strSQL2 = strSQL2 + " AND CP01||CP10 Not In ('FCP907', 'FCP913', 'FCP404', 'FG907', 'FG913', 'FG404') "
   StrSQL3 = StrSQL3 + " AND CP01||CP10 Not In ('FCP907', 'FCP913', 'FCP404', 'FG907', 'FG913', 'FG404') "
   StrSQL4 = StrSQL4 + " AND CP01||CP10 Not In ('FCP907', 'FCP913', 'FCP404', 'FG907', 'FG913', 'FG404') "
   strSQL5 = strSQL5 + " AND CP01||CP10 Not In ('FCP907', 'FCP913', 'FCP404', 'FG907', 'FG913', 'FG404') "
   'End
   If Len(frm050203.Text1(0)) <> 0 Then
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(frm050203.Text1(0), 1) & ") "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(frm050203.Text1(0), 2) & ") "
      StrSQL3 = StrSQL3 + " AND CP01 IN (" & SQLGrpStr(frm050203.Text1(0), 3) & ") "
      StrSQL4 = StrSQL4 + " AND CP01 IN (" & SQLGrpStr(frm050203.Text1(0), 4) & ") "
      strSQL5 = strSQL5 + " AND CP01 IN (" & SQLGrpStr(frm050203.Text1(0), 5) & ") "
      pub_QL05 = pub_QL05 & ";" & frm050203.Label7 & frm050203.Text1(0) 'Add By Sindy 2010/9/28
   End If
   StrSQL6 = ""
   If Len(Trim(frm050203.Text1(1))) <> 0 Then
      StrSQL6 = StrSQL6 + " AND CP27>=" & Val(ChangeTStringToWString(frm050203.Text1(1))) & " "
      'Add by Amy 2013/11/20 +if
      If FMPFFPstate = "ADD" Then
        strFMPFFP3 = strFMPFFP3 + " And CP27>=" & Val(ChangeTStringToWString(frm050203.Text1(1))) & " "
      End If
   End If
   If Len(Trim(frm050203.Text1(2))) <> 0 Then
      StrSQL6 = StrSQL6 & " AND CP27<=" & Val(ChangeTStringToWString(frm050203.Text1(2))) & " "
      'Add by Amy 2013/11/20 +if
      If FMPFFPstate = "ADD" Then
        strFMPFFP3 = strFMPFFP3 + " And CP27<=" & Val(ChangeTStringToWString(frm050203.Text1(2))) & " "
      End If
   End If
   If Len(Trim(frm050203.Text1(1))) <> 0 Or Len(Trim(frm050203.Text1(2))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm050203.Label1 & frm050203.Text1(1) & "-" & frm050203.Text1(2) 'Add By Sindy 2010/9/28
   End If
   '93.10.20 MODIFY BY SONIA
   'StrSQL6 = StrSQL6 & " and cp20 is null "
   'Modify by Amy 2017/08/17 原:CP57 IS NULL
   StrSQL6 = StrSQL6 & " and cp20 is null AND CP159=0 "
   'Add by Amy 2013/11/20 +if
   If FMPFFPstate = "ADD" Then
        strFMPFFP3 = strFMPFFP3 + " And cp20 is null And CP159=0 "
   End If
   'end 2017/08/17
   '93.10.20 END
   '93.6.16 ADD BY SONIA
   If frm050203.Text1(11).Text = "1" Then
      StrSQL6 = StrSQL6 & " AND CP16 > 0 "
      'Add by Amy 2013/11/20 +if
      If FMPFFPstate = "ADD" Then
        strFMPFFP3 = strFMPFFP3 + " And CP16 > 0 "
      End If
      pub_QL05 = pub_QL05 & ";" & Left(frm050203.Label4(3), 5) & "收據 / 請款單" 'Add By Sindy 2010/9/28
   End If
   '93.6.16 END
   'Add By Cheng 2002/12/17
   'A類不稽核 403, 411, 418, 419, 901, 902, 908, B類只稽核 907, 913, C類都不稽核
   '2010/5/26 MODIFY BY SONIA 此條件應限制FCP,FG
   'StrSQL6 = StrSQL6 & " And ((CP09<'B' AND (CP10<>'403' AND CP10<>'411' AND CP10<>'418' AND CP10<>'419' AND CP10<>'901' AND CP10<>'902' AND CP10<>'908')) OR ((CP09>='B' AND CP09<'C') AND (CP10='907' OR CP10='913'))) "
   If InStr(1, frm050203.Text1(0), "FCP") > 0 Or InStr(1, frm050203.Text1(0), "FG") > 0 Then
      'modify by sonia 2020/9/2 FCP-055795(BA6025643),FCP-017515(B96046390),FCP-052056(BA7031897)DAVID同意應檢查
      'StrSQL6 = StrSQL6 & " And ((CP09<'B' AND (CP10<>'403' AND CP10<>'411' AND CP10<>'418' AND CP10<>'419' AND CP10<>'901' AND CP10<>'902' AND CP10<>'908')) OR ((CP09>='B' AND CP09<'C') AND (CP10='907' OR CP10='913'))) "
      StrSQL6 = StrSQL6 & " And ((CP09<'B' AND (CP10<>'403' AND CP10<>'411' AND CP10<>'418' AND CP10<>'419' AND CP10<>'901' AND CP10<>'902' AND CP10<>'908')) OR (CP09>='B' AND CP09<'C' AND CP10<>'907' AND CP10<>'913')) "
   End If
   '2010/5/26 END
     
   If Len(Trim(frm050203.Text1(3))) <> 0 Then
       strSQL1 = strSQL1 + " AND PA09>='" & frm050203.Text1(3) & "' "
       strSQL2 = strSQL2 + " AND TM10>='" & frm050203.Text1(3) & "' "
       StrSQL3 = StrSQL3 + " AND LC15>='" & frm050203.Text1(3) & "' "
       'add by nickc 2007/01/11 台灣才出來
       If Trim(frm050203.Text1(3)) = "000" Or Trim(frm050203.Text1(3)) = "" Then
           StrSQL4 = StrSQL4 + " AND 1=1 "
       Else
           StrSQL4 = StrSQL4 + " AND 1<>1 "
       End If
       strSQL5 = strSQL5 + " AND SP09>='" & frm050203.Text1(3) & "' "
       'Add by Amy 2013/11/20 +if
       If FMPFFPstate = "ADD" Then
            strFMPFFP2 = strFMPFFP2 + " And SP09>='" & frm050203.Text1(3) & "' "
       End If
   End If
   If Len(Trim(frm050203.Text1(4))) <> 0 Then
       strSQL1 = strSQL1 + " AND PA09<='" & frm050203.Text1(4) & "' "
       strSQL2 = strSQL2 + " AND TM10<='" & frm050203.Text1(4) & "' "
       StrSQL3 = StrSQL3 + " AND LC15<='" & frm050203.Text1(4) & "' "
       'add by nickc 2007/01/11 台灣才出來
       If Trim(frm050203.Text1(4)) = "000" Or Trim(frm050203.Text1(4)) = "" Then
           StrSQL4 = StrSQL4 + " AND 1=1 "
       Else
           StrSQL4 = StrSQL4 + " AND 1<>1 "
       End If
       strSQL5 = strSQL5 + " AND SP09<='" & frm050203.Text1(4) & "' "
        'Add by Amy 2013/11/20 +if
       If FMPFFPstate = "ADD" Then
            strFMPFFP2 = strFMPFFP2 + " And SP09<='" & frm050203.Text1(4) & "' "
       End If
   End If
   If Len(Trim(frm050203.Text1(3))) <> 0 Or Len(Trim(frm050203.Text1(4))) <> 0 Then
      pub_QL05 = pub_QL05 & ";" & frm050203.Label2 & Trim(frm050203.Text1(3)) & "-" & Trim(frm050203.Text1(4)) 'Add By Sindy 2010/9/28
   End If
   'edit by nickc 2007/01/10
   If Len(Trim(frm050203.Text1(5))) <> 0 And Len(Trim(frm050203.Text1(7))) <> 0 Then
      '2013/11/18 modify by sonia 畫面條件改FC代理人,故此取消CP44條件
      'strSQL1 = strSQL1 + " and ((pa75>='" & GetNewFagent(frm050203.Text1(5)) & "' and pa75<='" & GetNewFagent(frm050203.Text1(7)) & "') or (cp44>='" & GetNewFagent(frm050203.Text1(5)) & "' and cp44<='" & GetNewFagent(frm050203.Text1(7)) & "')) "
      'strSQL2 = strSQL2 + " and ((tm44>='" & GetNewFagent(frm050203.Text1(5)) & "' and tm44<='" & GetNewFagent(frm050203.Text1(7)) & "') or (cp44>='" & GetNewFagent(frm050203.Text1(5)) & "' and cp44<='" & GetNewFagent(frm050203.Text1(7)) & "')) "
      'StrSQL3 = StrSQL3 + " and ((lc22>='" & GetNewFagent(frm050203.Text1(5)) & "' and lc22<='" & GetNewFagent(frm050203.Text1(7)) & "') or (cp44>='" & GetNewFagent(frm050203.Text1(5)) & "' and cp44<='" & GetNewFagent(frm050203.Text1(7)) & "')) "
      'Modify by Amy 2017/08/17 +(
      strSQL1 = strSQL1 + " and (pa75>='" & GetNewFagent(frm050203.Text1(5)) & "' and pa75<='" & GetNewFagent(frm050203.Text1(7)) & "') "
      strSQL2 = strSQL2 + " and (tm44>='" & GetNewFagent(frm050203.Text1(5)) & "' and tm44<='" & GetNewFagent(frm050203.Text1(7)) & "') "
      StrSQL3 = StrSQL3 + " and (lc22>='" & GetNewFagent(frm050203.Text1(5)) & "' and lc22<='" & GetNewFagent(frm050203.Text1(7)) & "') "
      '2013/11/18 END
      If Trim(frm050203.Text1(5)) = "" And Trim(frm050203.Text1(7)) = "" Then
         StrSQL4 = StrSQL4 & " and 1=1 "
      Else
         StrSQL4 = StrSQL4 & " and 1<>1 "
      End If
      '2013/11/18 modify by sonia 畫面條件改FC代理人,故此取消CP44條件
      'strSQL1 = strSQL1 + " and ((pa75>='" & GetNewFagent(frm050203.Text1(5
      'strSQL5 = strSQL5 + " and ((sp26>='" & GetNewFagent(frm050203.Text1(5)) & "' and sp26<='" & GetNewFagent(frm050203.Text1(7)) & "') or (cp44>='" & GetNewFagent(frm050203.Text1(5)) & "' and cp44<='" & GetNewFagent(frm050203.Text1(7)) & "')) "
      'Modify by Amy 2017/08/17 +(
      strSQL5 = strSQL5 + " and (sp26>='" & GetNewFagent(frm050203.Text1(5)) & "' and sp26<='" & GetNewFagent(frm050203.Text1(7)) & "') "
      'Add by Amy 2013/11/20 +if
       If FMPFFPstate = "ADD" Then
            'Modify by Amy 2017/08/17 +( 及strFMPFFP1
            strFMPFFP1 = strFMPFFP1 + " and (pa75>='" & GetNewFagent(frm050203.Text1(5)) & "' and pa75<='" & GetNewFagent(frm050203.Text1(7)) & "') "
            strFMPFFP2 = strFMPFFP2 + " and (sp26>='" & GetNewFagent(frm050203.Text1(5)) & "' and sp26<='" & GetNewFagent(frm050203.Text1(7)) & "') "
       End If
      '2013/11/18 END
      pub_QL05 = pub_QL05 & ";" & frm050203.Label4(0) & Trim(frm050203.Text1(5)) & "-" & Trim(frm050203.Text1(7)) 'Add By Sindy 2010/9/28
   End If
   'edit by nickc 2007/01/10
   
   If Len(Trim(frm050203.Text1(6))) <> 0 And Len(Trim(frm050203.Text1(8))) <> 0 Then
      'Modify By Sindy 2011/2/18 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
      strSQL1 = strSQL1 + " and ((pa26>='" & GetNewFagent(frm050203.Text1(6)) & "' and pa26<='" & GetNewFagent(frm050203.Text1(8)) & "') or (pa27>='" & GetNewFagent(frm050203.Text1(6)) & "' and pa27<='" & GetNewFagent(frm050203.Text1(8)) & "') or (pa28>='" & GetNewFagent(frm050203.Text1(6)) & "' and pa28<='" & GetNewFagent(frm050203.Text1(8)) & "') or (pa29>='" & GetNewFagent(frm050203.Text1(6)) & "' and pa29<='" & GetNewFagent(frm050203.Text1(8)) & "') or (pa30>='" & GetNewFagent(frm050203.Text1(6)) & "' and pa30<='" & GetNewFagent(frm050203.Text1(8)) & "')) "
      strSQL2 = strSQL2 + " and ((tm23>='" & GetNewFagent(frm050203.Text1(6)) & "' and tm23<='" & GetNewFagent(frm050203.Text1(8)) & "') or (tm78>='" & GetNewFagent(frm050203.Text1(6)) & "' and tm78<='" & GetNewFagent(frm050203.Text1(8)) & "') or (tm79>='" & GetNewFagent(frm050203.Text1(6)) & "' and tm79<='" & GetNewFagent(frm050203.Text1(8)) & "') or (tm80>='" & GetNewFagent(frm050203.Text1(6)) & "' and tm80<='" & GetNewFagent(frm050203.Text1(8)) & "') or (tm81>='" & GetNewFagent(frm050203.Text1(6)) & "' and tm81<='" & GetNewFagent(frm050203.Text1(8)) & "')) "
      StrSQL3 = StrSQL3 + " and ((lc11>='" & GetNewFagent(frm050203.Text1(6)) & "' and lc11<='" & GetNewFagent(frm050203.Text1(8)) & "') or (lc43>='" & GetNewFagent(frm050203.Text1(6)) & "' and lc43<='" & GetNewFagent(frm050203.Text1(8)) & "') or (lc44>='" & GetNewFagent(frm050203.Text1(6)) & "' and lc44<='" & GetNewFagent(frm050203.Text1(8)) & "') or (lc45>='" & GetNewFagent(frm050203.Text1(6)) & "' and lc45<='" & GetNewFagent(frm050203.Text1(8)) & "') or (lc46>='" & GetNewFagent(frm050203.Text1(6)) & "' and lc46<='" & GetNewFagent(frm050203.Text1(8)) & "')) "
      StrSQL4 = StrSQL4 + " and ((hc05>='" & GetNewFagent(frm050203.Text1(6)) & "' and hc05<='" & GetNewFagent(frm050203.Text1(8)) & "') or (hc24>='" & GetNewFagent(frm050203.Text1(6)) & "' and hc24<='" & GetNewFagent(frm050203.Text1(8)) & "') or (hc25>='" & GetNewFagent(frm050203.Text1(6)) & "' and hc25<='" & GetNewFagent(frm050203.Text1(8)) & "') or (hc26>='" & GetNewFagent(frm050203.Text1(6)) & "' and hc26<='" & GetNewFagent(frm050203.Text1(8)) & "') or (hc27>='" & GetNewFagent(frm050203.Text1(6)) & "' and hc27<='" & GetNewFagent(frm050203.Text1(8)) & "')) "
      strSQL5 = strSQL5 + " and ((sp08>='" & GetNewFagent(frm050203.Text1(6)) & "' and sp08<='" & GetNewFagent(frm050203.Text1(8)) & "') or (sp58>='" & GetNewFagent(frm050203.Text1(6)) & "' and sp58<='" & GetNewFagent(frm050203.Text1(8)) & "') or (sp59>='" & GetNewFagent(frm050203.Text1(6)) & "' and sp59<='" & GetNewFagent(frm050203.Text1(8)) & "') or (sp65>='" & GetNewFagent(frm050203.Text1(6)) & "' and sp65<='" & GetNewFagent(frm050203.Text1(8)) & "') or (sp66>='" & GetNewFagent(frm050203.Text1(6)) & "' and sp66<='" & GetNewFagent(frm050203.Text1(8)) & "')) "
      pub_QL05 = pub_QL05 & ";" & frm050203.Label4(1) & Trim(frm050203.Text1(6)) & "-" & Trim(frm050203.Text1(8)) 'Add By Sindy 2010/9/28
   End If
   
   'Add by Morgan 2003/12/04
   If Len(Trim(frm050203.Text1(10))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP10='" & frm050203.Text1(10) & "' "
      'add by  nickc 2007/01/11
      strSQL2 = strSQL2 + " AND CP10='" & frm050203.Text1(10) & "' "
      StrSQL3 = StrSQL3 + " AND CP10='" & frm050203.Text1(10) & "' "
      StrSQL4 = StrSQL4 + " AND CP10='" & frm050203.Text1(10) & "' "
      strSQL5 = strSQL5 + " AND CP10='" & frm050203.Text1(10) & "' "
      pub_QL05 = pub_QL05 & ";" & frm050203.Label3 & Trim(frm050203.Text1(10)) 'Add By Sindy 2010/9/28
   End If
   'Add by Amy 2016/12/07 +業務區
   If Len(Trim(frm050203.Text1(12))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP12>='" & frm050203.Text1(12) & "' "
      strSQL2 = strSQL2 + " AND CP12>='" & frm050203.Text1(12) & "' "
      StrSQL3 = StrSQL3 + " AND CP12>='" & frm050203.Text1(12) & "' "
      StrSQL4 = StrSQL4 + " AND CP12>='" & frm050203.Text1(12) & "' "
      strSQL5 = strSQL5 + " AND CP12>='" & frm050203.Text1(12) & "' "
      pub_QL05 = pub_QL05 & ";" & frm050203.Label10 & Trim(frm050203.Text1(12))
      End If
   If Len(Trim(frm050203.Text1(13))) <> 0 Then
     strSQL1 = strSQL1 + " AND CP12<='" & frm050203.Text1(13) & "' "
      strSQL2 = strSQL2 + " AND CP12<='" & frm050203.Text1(13) & "' "
      StrSQL3 = StrSQL3 + " AND CP12<='" & frm050203.Text1(13) & "' "
      StrSQL4 = StrSQL4 + " AND CP12<='" & frm050203.Text1(13) & "' "
      strSQL5 = strSQL5 + " AND CP12<='" & frm050203.Text1(13) & "' "
      pub_QL05 = pub_QL05 & ";" & frm050203.Label10 & "-" & Trim(frm050203.Text1(13))
   End If
   'end 2016/12/07
   
   cnnConnection.Execute "DELETE FROM R050203 WHERE ID='" & strUserNum & "' "
   
   'Modify By Cheng 2002/04/26
   '若已閉卷, 則在本所案號後加"*"號
   '2008/11/10 modify by sonia FCP,FG智權人員抓國家檔之NA51,不抓申請人國籍改抓代理人國籍
   'Modify by Amy 2013/11/20 不抓cp44改抓申請人 R01007=申請人國籍 R01008=代理人國籍 +CP12/CP10/PA150 or SP79/CP09
   'Modified by Lydia 2014/11/14 NA51 = decode(pa75," & midstr & ",na51)
   If InStr(1, frm050203.Text1(0), "FCP") > 0 Or InStr(1, frm050203.Text1(0), "FG") > 0 Then
                          'strSql = "SELECT CP14,NA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,NVL(F1.FA10,F2.FA10),NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,CP44,'" & strUserNum & "',CP01,CP12,CP10,PA150 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL1 & StrSQL6
      'strSql = strSql + " UNION ALL SELECT CP14,NA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),nvl(TM05,nvl(TM06,TM07)),NVL(decode(tm10,'000',CPM03,cpm04),CP10),TM10,NVL(F1.FA10,F2.FA10),NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",TM44,CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM TRADEMARK,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP,NATION WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND DECODE(SUBSTr(TM44,9,1),'','0',SUBSTR(TM44,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL2 & StrSQL6
      'strSql = strSql + " UNION ALL SELECT CP14,NA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),nvl(LC05,nvl(LC06,LC07)),NVL(decode(lc15,'000',cpm03,cpm04),CP10),LC15,NVL(F1.FA10,F2.FA10),NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",LC22,CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM LAWCASE,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP,NATION WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND DECODE(SUBSTr(LC22,9,1),'','0',SUBSTR(LC22,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & StrSQL3 & StrSQL6
      'strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,cpm03,'000',CU10,F2.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",' ',CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM HIRECASE,CASEPROGRESS,CUSTOMER,FAGENT F2,CASEPROPERTYMAP WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP60 IS NULL AND CP27 IS NOT NULL AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & StrSQL6
      'strSql = strSql + " UNION ALL SELECT CP14,NA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,NVL(F1.FA10,F2.FA10),NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL5 & StrSQL6
      'Modified by Lydia 2016/02/03
'                          strSql = "SELECT CP14,decode(pa75,'Y51333010','" & midStr & "',na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,PA26,'" & strUserNum & "',CP01,CP12,CP10,PA150,CP09 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL1 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,decode(tm44,'Y51333010','" & midStr & "',na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),nvl(TM05,nvl(TM06,TM07)),NVL(decode(tm10,'000',CPM03,cpm04),CP10),TM10,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",TM44,TM23,'" & strUserNum & "',CP01,CP12,CP10,'',CP09  FROM TRADEMARK,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND DECODE(SUBSTr(TM44,9,1),'','0',SUBSTR(TM44,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL2 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,decode(lc22,'Y51333010','" & midStr & "',na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),nvl(LC05,nvl(LC06,LC07)),NVL(decode(lc15,'000',cpm03,cpm04),CP10),LC15,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",LC22,LC11,'" & strUserNum & "',CP01,CP12,CP10,'',CP09  FROM LAWCASE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND DECODE(SUBSTr(LC22,9,1),'','0',SUBSTR(LC22,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & StrSQL3 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,cpm03,'000',CU10,''," & SQLDate("CP05") & "," & SQLDate("CP27") & ",' ','','" & strUserNum & "',CP01,CP12,CP10,'',CP09  FROM HIRECASE,CASEPROGRESS,CUSTOMER,CASEPROPERTYMAP WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP60 IS NULL AND CP27 IS NOT NULL AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,decode(sp26,'Y51333010','" & midStr & "',na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09  FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL5 & StrSQL6
      'Add by Amy 2016/12/07 +CP84
                          strSql = "SELECT CP14,decode(pa75," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,PA26,'" & strUserNum & "',CP01,CP12,CP10,PA150,CP09,CP84,CP27 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL1 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,decode(tm44," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),nvl(TM05,nvl(TM06,TM07)),NVL(decode(tm10,'000',CPM03,cpm04),CP10),TM10,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",TM44,TM23,'" & strUserNum & "',CP01,CP12,CP10,'',CP09,CP84,CP27  FROM TRADEMARK,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND DECODE(SUBSTr(TM44,9,1),'','0',SUBSTR(TM44,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL2 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,decode(lc22," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),nvl(LC05,nvl(LC06,LC07)),NVL(decode(lc15,'000',cpm03,cpm04),CP10),LC15,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",LC22,LC11,'" & strUserNum & "',CP01,CP12,CP10,'',CP09,CP84,CP27  FROM LAWCASE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND DECODE(SUBSTr(LC22,9,1),'','0',SUBSTR(LC22,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & StrSQL3 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,cpm03,'000',CU10,''," & SQLDate("CP05") & "," & SQLDate("CP27") & ",' ','','" & strUserNum & "',CP01,CP12,CP10,'',CP09,CP84,CP27  FROM HIRECASE,CASEPROGRESS,CUSTOMER,CASEPROPERTYMAP WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP60 IS NULL AND CP27 IS NOT NULL AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,decode(sp26," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09,CP84,CP27  FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strSQL5 & StrSQL6
  
      'Add by Amy 2013/11/20 +FMP、FFP
      If FMPFFPstate = "ADD" Then
      'Modified by Lydia 2016/02/03
'            strSql = strSql + " Union All Select CP14,decode(pa75,'Y51333010','" & midStr & "',na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),Decode(CP61,null,'',CP61||'(帳單)'),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,PA26,'" & strUserNum & "',CP01,CP12,CP10,PA150,CP09 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strFMPFFP1 & strFMPFFP3
'            strSql = strSql + " Union All Select CP14,decode(sp26,'Y51333010','" & midStr & "',na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),Decode(CP61,null,'',CP61||'(帳單)'),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09 FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strFMPFFP2 & strFMPFFP3
            'Add by Amy 2016/12/07 +CP84,CP27
            strSql = strSql + " Union All Select CP14,decode(pa75," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),Decode(CP61,null,'',CP61||'(帳單)'),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,PA26,'" & strUserNum & "',CP01,CP12,CP10,PA150,CP09,CP84,CP27 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strFMPFFP1 & strFMPFFP3
            strSql = strSql + " Union All Select CP14,decode(sp26," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),Decode(CP61,null,'',CP61||'(帳單)'),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09,CP84,CP27 FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP,NATION WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) AND F1.FA10=NA01(+) " & strFMPFFP2 & strFMPFFP3
      
      End If
   Else
'                          strSql = "SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,CP44,'" & strUserNum & "',CP01,CP12,CP10,PA150 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),nvl(TM05,nvl(TM06,TM07)),NVL(decode(tm10,'000',CPM03,cpm04),CP10),TM10,CU10,NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",TM44,CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM TRADEMARK,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND DECODE(SUBSTr(TM44,9,1),'','0',SUBSTR(TM44,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),nvl(LC05,nvl(LC06,LC07)),NVL(decode(lc15,'000',cpm03,cpm04),CP10),LC15,CU10,NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",LC22,CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM LAWCASE,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND DECODE(SUBSTr(LC22,9,1),'','0',SUBSTR(LC22,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,cpm03,'000',CU10,F2.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",' ',CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM HIRECASE,CASEPROGRESS,CUSTOMER,FAGENT F2,CASEPROPERTYMAP WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP60 IS NULL AND CP27 IS NOT NULL AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & StrSQL6
'      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,NVL(F1.FA10,F2.FA10)," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,CP44,'" & strUserNum & "',CP01,CP12,CP10,''  FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,FAGENT F2,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND SUBSTR(CP44,1,8)=F2.FA01(+) AND DECODE(SUBSTR(CP44,9,1),'','0',SUBSTR(CP44,9,1)) = F2.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & StrSQL6
      'Add by Amy 2016/12/07 +CP84,CP27
                          strSql = "SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,PA26,'" & strUserNum & "',CP01,CP12,CP10,PA150,CP09,CP84,CP27 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL1 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(TM29,'Y','＊',''),nvl(TM05,nvl(TM06,TM07)),NVL(decode(tm10,'000',CPM03,cpm04),CP10),TM10,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",TM44,TM23,'" & strUserNum & "',CP01,CP12,CP10,'',CP09,CP84,CP27 FROM TRADEMARK,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(TM23,1,8)=CU01(+) AND DECODE(SUBSTR(TM23,9,1),'','0',SUBSTR(TM23,9,1))=CU02(+) AND SUBSTR(TM44,1,8)=F1.FA01(+) AND DECODE(SUBSTr(TM44,9,1),'','0',SUBSTR(TM44,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL2 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(LC08,'Y','＊',''),nvl(LC05,nvl(LC06,LC07)),NVL(decode(lc15,'000',cpm03,cpm04),CP10),LC15,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",LC22,LC11,'" & strUserNum & "',CP01,CP12,CP10,'',CP09,CP84,CP27  FROM LAWCASE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=LC01(+) AND CP02=LC02(+) AND CP03=LC03(+) AND CP04=LC04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(LC11,1,8)=CU01(+) AND DECODE(SUBSTR(LC11,9,1),'','0',SUBSTR(LC11,9,1))=CU02(+) AND SUBSTR(LC22,1,8)=F1.FA01(+) AND DECODE(SUBSTr(LC22,9,1),'','0',SUBSTR(LC22,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL3 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(HC09,'Y','＊',''),HC06,cpm03,'000',CU10,''," & SQLDate("CP05") & "," & SQLDate("CP27") & ",' ','','" & strUserNum & "',CP01,CP12,CP10,'',CP09,CP84,CP27  FROM HIRECASE,CASEPROGRESS,CUSTOMER,CASEPROPERTYMAP WHERE CP01=HC01(+) AND CP02=HC02(+) AND CP03=HC03(+) AND CP04=HC04(+) AND CP60 IS NULL AND CP27 IS NOT NULL AND SUBSTR(HC05,1,8)=CU01(+) AND DECODE(SUBSTR(HC05,9,1),'','0',SUBSTR(HC05,9,1))=CU02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & StrSQL4 & StrSQL6
      strSql = strSql + " UNION ALL SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09,CP84,CP27  FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND CP27 IS NOT NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & strSQL5 & StrSQL6
   End If
   
   CheckOC
   'Add by Amy 2016/12/07 +發文規費/帳單金額
   'Modified by Lydia 2018/02/22 + lngQ
   cnnConnection.Execute "INSERT INTO R050203 (R01001,R01002,R01003,R01004,R01005,R01006,R01007,R01008,R01009,R01010,R01011,R01012,ID,R01013,R01014,R01015,R01016,R01017,R01018,R01020) " & strSql, lngQ
   
   'Added by Lydia 2018/02/22 判斷暫存檔無資料就結束
   If lngQ = 0 Then
        GoTo JumpEnd
   End If
   'end 2018/02/22
   
   'Add by Amy 2013/11/20 +系別有FCP
   If InStr(1, frm050203.Text1(0), "FCP") > 0 Then
        'Add by Amy 2016/12/07
        '案件性質101,102,103,106,125，檢查該案之201,209,210,235發文日(若同時存在 201or 209 or 210 or 235 抓收文日較早的發文日比較 ex:FCP-048202)
        strExc(0) = "Select R01017,R01003 From R050203 Where ID='" & strUserNum & "' And R01015 In ('101','102','103','106','125') And R01013='FCP' " & _
                          "Order by R01003"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            Do While Not RsTemp.EOF
                strExc(1) = "Select CP05,Nvl(CP27,0) as CP27 From CaseProgress Where CP01='FCP' And CP02='" & Mid(RsTemp.Fields("R01003"), 5, 6) & "' " & _
                                "And CP03='" & Mid(RsTemp.Fields("R01003"), 12, 1) & "' And CP04='" & Mid(RsTemp.Fields("R01003"), 14, 2) & "' " & _
                                "And cp10 in ('201','209','210', '235') Order by CP05 asc"
                intI = 1 'Added by Lydia 2018/02/22
                Set adoRecordset = ClsLawReadRstMsg(intI, strExc(1))
                If intI = 1 Then
                    With adoRecordset
                        If Val(.Fields("CP27")) = 0 Or Val(.Fields("CP27")) > Val(frm050203.Text1(2)) + 19110000 Then
                            '若未發文或其發文日>畫面發文止日條件，則101,102,103,106不列
                            strExc(2) = "Delete From R050203 Where ID='" & strUserNum & "' And R01017='" & RsTemp.Fields("R01017") & "' "
                            cnnConnection.Execute strExc(2)
                        Else
                            '更新為201,209,210,235發文日以計算未請款月數
                            strExc(2) = "Update R050203 Set R01020='" & .Fields("CP27") & "' Where  ID='" & strUserNum & "' And R01003='" & RsTemp.Fields("R01003") & "' "
                            cnnConnection.Execute strExc(2)
                        End If
                    End With
                End If
                RsTemp.MoveNext
            Loop
        End If
        If adoRecordset.State = adStateOpen Then adoRecordset.Close
        If RsTemp.State = adStateOpen Then RsTemp.Close
        '案件性質416,202,203，檢查該案之201,209,210,235之資料(若同時存在 201or 209 or 210 or 235 抓收文日較早的發文日比較)
        strExc(0) = "Select R01017,R01003,R01009 as CP05,Nvl(R01010,0) as R01010 From R050203 Where ID='" & strUserNum & "' And R01015 In ('416','202','203')  And R01013='FCP' " & _
                          "Order by R01003"
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
        If intI = 1 Then
            Do While Not RsTemp.EOF
                strExc(1) = "Select CP05,Nvl(CP27,0) as CP27 From CaseProgress Where CP01='FCP' And CP02='" & Mid(RsTemp.Fields("R01003"), 5, 6) & "' " & _
                                "And CP03='" & Mid(RsTemp.Fields("R01003"), 12, 1) & "' And CP04='" & Mid(RsTemp.Fields("R01003"), 14, 2) & "' " & _
                                "And cp10 in ('201','209','210', '235') Order by CP05 asc"
                intI = 1 'Added by Lydia 2018/02/22
                Set adoRecordset = ClsLawReadRstMsg(intI, strExc(1))
                If intI = 1 Then
                    With adoRecordset
                        If Val(.Fields("CP27")) = 0 Or (Val(.Fields("CP27")) > Val(frm050203.Text1(2)) + 19110000) And .Fields("CP27") >= FCDate(RsTemp.Fields("CP05")) + 19110000 Then
                            '若201,209,210,235已發文且發文日大於416,202,203之收文日或其發文日>畫面發文止日條件，則416,202,203不列
                            strExc(2) = "Delete From R050203 Where ID='" & strUserNum & "' And R01017='" & RsTemp.Fields("R01017") & "' "
                            cnnConnection.Execute strExc(2)
                        ElseIf Val(.Fields("CP27")) > 0 And (Val(.Fields("CP27")) >= Val(FCDate(RsTemp.Fields("R01010"))) + 19110000) Then
                            '若201,209,210,235發文日大於或等於416,202,203之收文日，則同本所案號都以201,209,210,235之發文日計算未請款月數
                            strExc(2) = "Update R050203 Set R01020='" & .Fields("CP27") & "' Where  ID='" & strUserNum & "' And R01003='" & RsTemp.Fields("R01003") & "' "
                            cnnConnection.Execute strExc(2)
                        End If
                    End With
                End If
                RsTemp.MoveNext
            Loop
        End If
        If adoRecordset.State = adStateOpen Then adoRecordset.Close
        If RsTemp.State = adStateOpen Then RsTemp.Close
        'end 2016/12/07
        
        '有新申請案，需將該案其他程序未請款也帶出(不管是否發文)
        strExc(0) = "Select * From R050203 Where ID='" & strUserNum & "' And R01015 In (" & NewCasePtyList & ") Order by R01003 "
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
                strTp(2) = Pub_RplStr(RsTemp.Fields("R01003"))
                'Add by Amy 2013/12/02 +Union 案件性質為201,209,210,235
                Select Case SystemNumber(strTp(2), 1)
                    Case "CFP", "FCP", "P"   '專利
                        'Modify by Amy 2016/12/07 +CP84
                        'Modify by Amy 2017/01/04 CP44改抓PA26,Union 排除暫檔已存在
                        strExc(1) = "Insert Into R050203 (R01001,R01002,R01003,R01004,R01005,R01006,R01007,R01008,R01009,R01010,R01011,R01012,ID,R01013,R01014,R01015,R01016,R01017,R01018,R01020) " & _
                                         "SELECT CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,PA26,'" & strUserNum & "',CP01,CP12,CP10,PA150,CP09,CP84,CP27 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
                                         "And CP16 >0 And CP20 is null AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                                         "And PA01='" & SystemNumber(strTp(2), 1) & "' And PA02='" & SystemNumber(strTp(2), 2) & "' And PA03='" & SystemNumber(strTp(2), 3) & "' And PA04='" & SystemNumber(strTp(2), 4) & "' And CP10 Not In ( Select R01015 From R050203 Where R01003='" & RsTemp.Fields("R01003") & "' and ID='" & strUserNum & "' ) " & _
                                         "Union Select CP14,CP13,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),nvl(pa05,nvl(pa06,pa07)),NVL(decode(pa09,'000',CPM03,CPM04),cp10),PA09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",PA75,PA26,'" & strUserNum & "',CP01,CP12,CP10,PA150,CP09,CP84,CP27 FROM PATENT,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 " & _
                                         "And CP10 In (201,209,210,235) And decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND SUBSTR(PA26,1,8)=CU01(+) AND DECODE(SUBSTR(PA26,9,1),'','0',SUBSTR(PA26,9,1))=CU02(+) AND SUBSTR(PA75,1,8)=F1.FA01(+) AND DECODE(SUBSTr(PA75,9,1),'','0',SUBSTR(PA75,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                                         "And PA01='" & SystemNumber(strTp(2), 1) & "' And PA02='" & SystemNumber(strTp(2), 2) & "' And PA03='" & SystemNumber(strTp(2), 3) & "' And PA04='" & SystemNumber(strTp(2), 4) & "' And CP10 Not In ( Select R01015 From R050203 Where R01003='" & RsTemp.Fields("R01003") & "' and ID='" & strUserNum & "' ) "
                
                    Case "FG" '服務
                        'Modified by Lydia 2014/11/14 NA51 = decode(pa75," & midstr & ",na51)
                        'Modified by Lydia 2016/02/03
'                        strExc(1) = "Insert Into R050203 " & _
'                                         "SELECT CP14,decode(SP26," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09  FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
'                                         "And CP16 >0 And CP20 is null AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                                         "And SP01='" & SystemNumber(strTp(2), 1) & "' And SP02='" & SystemNumber(strTp(2), 2) & "' And SP03='" & SystemNumber(strTp(2), 3) & "' And SP04='" & SystemNumber(strTp(2), 4) & "' And CP10 Not In ( Select R01015 From R050203 Where R01003='" & RsTemp.Fields("R01003") & "' and ID='" & strUserNum & "' ) " & _
'                                         "Union Select CP14,decode(SP26," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09  FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
'                                         "And CP10 In (201,209,210,235) And decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
'                                         "And SP01='" & SystemNumber(strTp(2), 1) & "' And SP02='" & SystemNumber(strTp(2), 2) & "' And SP03='" & SystemNumber(strTp(2), 3) & "' And SP04='" & SystemNumber(strTp(2), 4) & "' "
                        'Modify by Amy 2016/12/07 +CP84
                        'Modify by Amy 2017/01/04 Union 排除暫檔已存在
                        strExc(1) = "Insert Into R050203 (R01001,R01002,R01003,R01004,R01005,R01006,R01007,R01008,R01009,R01010,R01011,R01012,ID,R01013,R01014,R01015,R01016,R01017,R01018,R01020) " & _
                                         "SELECT CP14,decode(SP26," & midStr & ",na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09,CP84,CP27 FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
                                         "And CP16 >0 And CP20 is null AND decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                                         "And SP01='" & SystemNumber(strTp(2), 1) & "' And SP02='" & SystemNumber(strTp(2), 2) & "' And SP03='" & SystemNumber(strTp(2), 3) & "' And SP04='" & SystemNumber(strTp(2), 4) & "' And CP10 Not In ( Select R01015 From R050203 Where R01003='" & RsTemp.Fields("R01003") & "' and ID='" & strUserNum & "' ) " & _
                                         "Union Select CP14,decode(SP26,'Y51333010','" & midStr & "',na51) nNA51,CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(SP15,'Y','＊',''),nvl(Sp05,nvl(Sp06,Sp07)),nvl(decode(sp09,'000',CPM03,CPM04),cp10),SP09,CU10,F1.FA10," & SQLDate("CP05") & "," & SQLDate("CP27") & ",SP26,SP08,'" & strUserNum & "',CP01,CP12,CP10,SP79,CP09,CP84,CP27 FROM SERVICEPRACTICE,CASEPROGRESS,CUSTOMER,FAGENT F1,CASEPROPERTYMAP WHERE CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) " & _
                                         "And CP10 In (201,209,210,235) And decode('" & frm050203.Text1(11).Text & "','1',CP60,cp61) IS NULL AND SUBSTR(SP08,1,8)=CU01(+) AND DECODE(SUBSTR(SP08,9,1),'','0',SUBSTR(SP08,9,1))=CU02(+) AND SUBSTR(SP26,1,8)=F1.FA01(+) AND DECODE(SUBSTr(SP26,9,1),'','0',SUBSTR(SP26,9,1))=F1.FA02(+) AND CP01=CPM01(+) AND CP10=CPM02(+) " & _
                                         "And SP01='" & SystemNumber(strTp(2), 1) & "' And SP02='" & SystemNumber(strTp(2), 2) & "' And SP03='" & SystemNumber(strTp(2), 3) & "' And SP04='" & SystemNumber(strTp(2), 4) & "' And CP10 Not In ( Select R01015 From R050203 Where R01003='" & RsTemp.Fields("R01003") & "' and ID='" & strUserNum & "' ) "

                End Select
                cnnConnection.Execute strExc(1)
                RsTemp.MoveNext
            Loop
         End If
         '重抓FCP承辦業務
         strExc(0) = "Select R01003,R01002,ID From R050203 Where ID='" & strUserNum & "' And R01013 ='FCP' Group by ID,R01003,R01002 Order by R01003"
          intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
                strTp(2) = Pub_RplStr(RsTemp.Fields("R01003"))
                strTp(0) = PUB_GetFCPSalesNo(SystemNumber(strTp(2), 1), SystemNumber(strTp(2), 2), SystemNumber(strTp(2), 3), SystemNumber(strTp(2), 4))
                If RsTemp.Fields("R01002") <> strTp(0) Then
                    strExc(1) = "Update R050203 Set R01002='" & strTp(0) & "' Where ID='" & strUserNum & "' And R01003='" & RsTemp.Fields("R01003") & "' "
                    cnnConnection.Execute strExc(1)
                End If
                RsTemp.MoveNext
            Loop
         End If
         'Add by Amy 2013/12/02 案件性質為201,209,210,235 重抓承辦人
         strExc(0) = "Select R01003,R01001,R01015,R01017,ID From R050203 Where ID='" & strUserNum & "' And R01015 In ('201','209','210','235') Group by ID,R01003,R01001,R01015,R01017 Order by R01003"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
                strTp(2) = Pub_RplStr(RsTemp.Fields("R01003"))
                strTp(0) = Get_Ep04(RsTemp.Fields("R01017"))
                If RsTemp.Fields("R01001") <> strTp(0) Then
                    strExc(1) = "Update R050203 Set R01001='" & strTp(0) & "' Where ID='" & strUserNum & "' And R01003='" & RsTemp.Fields("R01003") & "' And R01015='" & RsTemp.Fields("R01015") & "' "
                   cnnConnection.Execute strExc(1)
                End If
                RsTemp.MoveNext
            Loop
         End If
   ElseIf FMPFFPstate = "DEL" Then
        '剔除FMP、FFP (剔除CP12為F開頭之P、PS、CFP、CPS之案件)
        strFMPFFP1 = "Delete From R050203 Where ID='" & strUserNum & "' And R01013 In ('P','PS','CFP','CPS') And Substr(R01014,1,1)='F' "
        cnnConnection.Execute strFMPFFP1
   End If
   'Add by Amy 2016/12/07
   'FCP 同案號更新Max(R01020) ex:FCP-054582 於新申請案，需將該案其他程序未請款也帶出的資料其收文日會不一致
   strExc(0) = "Update R050203 O Set R01020=(Select Max(R01020) From R050203 Where ID='" & strUserNum & "' And R01013='FCP' And O.R01003=R01003 Group by R01003 ) " & _
                     "Where ID='" & strUserNum & "' And R01013='FCP' "
   cnnConnection.Execute strExc(0)
   '所有系統類別之延期，其承辦人抓相關總收文號之承辦人
   strExc(0) = "Update R050203 Set R01001=" & _
                    "(Select o.cp14 From CaseProgress a,CaseProgress o Where  R01017=a.cp09 And a.cp43=o.cp09) " & _
                    "Where ID='" & strUserNum & "' And ( (InStr(R01013,'P')>0 And R01015='404') OR (InStr(R01013,'T')>0 And R01015='303'))"
   cnnConnection.Execute strExc(0)
  '抓取帳單金額
  'Modify by Amy 2018/05/24 相同收文號金額加總
  strExc(0) = "Update R050203 Set R01019=" & _
                    "(Select Sum(Round(Nvl(Decode(a1901,null,axf04*a1g03,axf04*a1906),axf04*a2103),0)) From acc151,acc150,acc190,acc1g0,acc210 " & _
                    "Where R01017=axf02(+) And  axf01=a1501(+) And  axf01=a1902(+) And  a1512=A1G01(+) And  a1505=a2102(+) " & _
                    "And  a2101 = (Select Max(a2101) From acc210 Where a2102 = a1505 And  a2101 <= " & strSrvDate(2) & ") " & _
                    "Group by R01017 " & _
                    ") Where ID='" & strUserNum & "' "
   cnnConnection.Execute strExc(0)
   'end 2016/12/07
   '2013/11/18 modify by sonia 畫面條件改FC代理人,故此取消CP44
   'StrSQLa = "DECODE(SK1.SK03,0,NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)),DECODE(F1.FA05,NULL,NVL(F1.FA04,F1.FA06),F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)) "
   'StrSqlB = "DECODE(SK2.SK03,0,NVL(F2.FA04,DECODE(F2.FA05,NULL,F2.FA06,F2.FA05||' '||F2.FA63||' '||F2.FA64||' '||F2.FA65)),DECODE(F2.FA05,NULL,NVL(F2.FA04,F2.FA06),F2.FA05||' '||F2.FA63||' '||F2.FA64||' '||F2.FA65)) "
   ''2010/5/3 modify by sonia 代理人改為無FC代理人才抓CF代理人
   'strSql = "SELECT NVL(S1.ST02,R01001),NVL(S2.ST02,R01002),R01003,R01004,R01005,NVL(NVL(N1.NA03,N1.NA04),R01006),DECODE(R01006,'000',NVL(NVL(N2.NA03,N2.NA04),R01007),NVL(NVL(N3.NA03,N3.NA04),R01008)),R01009,R01010,DECODE(R01012,''," & StrSqlB & "," & StrSQLa & "),R01002 FROM R050203,NATION N1,NATION N2,NATION N3,FAGENT F1,FAGENT F2,STAFF S1,STAFF S2,SYSTEMKIND SK1,SYSTEMKIND SK2 WHERE id='" & strUserNum & "' and R01006=N1.NA01(+) AND R01007=N2.NA01(+) AND R01008=N3.NA01(+) AND SUBSTR(R01011,1,8)=F1.FA01(+) AND DECODE(SUBSTR(R01011,9,1),'','0',SUBSTR(R01011,9,1))=F1.FA02(+) AND SUBSTR(R01012,1,8)=F2.FA01(+) AND DECODE(SUBSTR(R01012,9,1),'','0',SUBSTR(R01012,9,1))=F2.FA02(+) AND R01001=S1.ST01(+) AND R01002=S2.ST01(+) AND R01013=SK1.SK01(+) AND R01013=SK2.SK01(+) "
   'Modify by Amy 2013/11/20 +if 改國籍代理人無資料抓申請人國籍 並抓PA150(R01016) R01013(系統別) ST03 CP09(R01017)
   'strSql = "SELECT NVL(S1.ST02,R01001),NVL(S2.ST02,R01002),R01003,R01004,R01005,NVL(NVL(N1.NA03,N1.NA04),R01006),DECODE(R01006,'000',NVL(NVL(N2.NA03,N2.NA04),R01007),NVL(NVL(N3.NA03,N3.NA04),R01008)),R01009,R01010,NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)),R01002 FROM R050203,NATION N1,NATION N2,NATION N3,FAGENT F1,STAFF S1,STAFF S2 WHERE id='" & strUserNum & "' and R01006=N1.NA01(+) AND R01007=N2.NA01(+) AND R01008=N3.NA01(+) AND SUBSTR(R01011,1,8)=F1.FA01(+) AND DECODE(SUBSTR(R01011,9,1),'','0',SUBSTR(R01011,9,1))=F1.FA02(+) AND R01001=S1.ST01(+) AND R01002=S2.ST01(+) "
    'Modify by Amy 2016/12/07 +發文規費/帳單金額,收文日改未請款逾月數=畫面發文日條件日年月-資料應計算之發文年月+1
    'strSql = "SELECT NVL(S1.ST02,R01001),NVL(S2.ST02,R01002),R01003,R01004,R01005,NVL(NVL(N1.NA03,N1.NA04),R01006),DECODE(R01011,null,NVL(NVL(N2.NA03,N2.NA04),R01007)||'(申)',NVL(NVL(N3.NA03,N3.NA04),R01008)),R01009,R01010,NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)),R01002,R01016,R01013,R01017 FROM R050203,NATION N1,NATION N2,NATION N3,FAGENT F1,STAFF S1,STAFF S2 " & _
                    "WHERE id='" & strUserNum & "' and R01006=N1.NA01(+) AND R01007=N2.NA01(+) AND R01008=N3.NA01(+) AND SUBSTR(R01011,1,8)=F1.FA01(+) AND DECODE(SUBSTR(R01011,9,1),'','0',SUBSTR(R01011,9,1))=F1.FA02(+) AND R01001=S1.ST01(+) AND R01002=S2.ST01(+) "
    strSql = Left(Val(frm050203.Text1(2)) + 19110000, 6) & "01"
    'Modified by Lydia 2018/03/22 單據類別:帳單的R01020有可能為空白 +Modified by Lydia 2018/03/01
    'strSql = "SELECT NVL(S1.ST02,R01001),NVL(S2.ST02,R01002),R01003,R01004,R01005,NVL(NVL(N1.NA03,N1.NA04),R01006),DECODE(R01011,null,NVL(NVL(N2.NA03,N2.NA04),R01007)||'(申)',NVL(NVL(N3.NA03,N3.NA04),R01008)),Months_Between(to_Date('" & strSql & "','YYYYMMDD'),to_Date(SubStr(R01020,1,6)||'01','YYYYMMDD'))+1,R01010,NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)),R01002,R01016,R01013,R01017,Nvl(R01018,0),Nvl(R01019,0) FROM R050203,NATION N1,NATION N2,NATION N3,FAGENT F1,STAFF S1,STAFF S2 " & _
                    "WHERE id='" & strUserNum & "' and R01006=N1.NA01(+) AND R01007=N2.NA01(+) AND R01008=N3.NA01(+) AND SUBSTR(R01011,1,8)=F1.FA01(+) AND DECODE(SUBSTR(R01011,9,1),'','0',SUBSTR(R01011,9,1))=F1.FA02(+) AND R01001=S1.ST01(+) AND R01002=S2.ST01(+) "
    ''end 2013/11/20
    ''2013/11/18 END
    strSql = "SELECT NVL(S1.ST02,R01001) A01,NVL(S2.ST02,R01002) A02,R01003,R01004,R01005,NVL(NVL(N1.NA03,N1.NA04),R01006),DECODE(R01011,null,NVL(NVL(N2.NA03,N2.NA04),R01007)||'(申)',NVL(NVL(N3.NA03,N3.NA04),R01008))," & _
                 " Months_Between(to_Date('" & strSql & "','YYYYMMDD'),to_Date(SubStr(R01020,1,6)||'01','YYYYMMDD'))+1,R01010,NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)),R01002,R01016,R01013,R01017,Nvl(R01018,0),Nvl(R01019,0),R01001,S2.ST03 " & _
                 "FROM R050203,NATION N1,NATION N2,NATION N3,FAGENT F1,STAFF S1,STAFF S2 " & _
                 "WHERE id='" & strUserNum & "' and R01006=N1.NA01(+) AND R01007=N2.NA01(+) AND R01008=N3.NA01(+) AND SUBSTR(R01011,1,8)=F1.FA01(+) AND DECODE(SUBSTR(R01011,9,1),'','0',SUBSTR(R01011,9,1))=F1.FA02(+) " & _
                 "AND R01001=S1.ST01(+) AND R01002=S2.ST01(+) AND NVL(R01020,0) > 0  "
    strSql = strSql & "UNION ALL SELECT NVL(S1.ST02,R01001) A01,NVL(S2.ST02,R01002) A02,R01003,R01004,R01005,NVL(NVL(N1.NA03,N1.NA04),R01006),DECODE(R01011,null,NVL(NVL(N2.NA03,N2.NA04),R01007)||'(申)',NVL(NVL(N3.NA03,N3.NA04),R01008))," & _
                 " 0,R01010,NVL(F1.FA04,DECODE(F1.FA05,NULL,F1.FA06,F1.FA05||' '||F1.FA63||' '||F1.FA64||' '||F1.FA65)),R01002,R01016,R01013,R01017,Nvl(R01018,0),Nvl(R01019,0),R01001,S2.ST03 " & _
                 "FROM R050203,NATION N1,NATION N2,NATION N3,FAGENT F1,STAFF S1,STAFF S2 " & _
                  "WHERE id='" & strUserNum & "' and R01006=N1.NA01(+) AND R01007=N2.NA01(+) AND R01008=N3.NA01(+) AND SUBSTR(R01011,1,8)=F1.FA01(+) AND DECODE(SUBSTR(R01011,9,1),'','0',SUBSTR(R01011,9,1))=F1.FA02(+) " & _
                  "AND R01001=S1.ST01(+) AND R01002=S2.ST01(+) AND NVL(R01020,0) = 0 "
    'end 2018/02/22
    
   'Add by Amy 2013/11/20
   If InStr(1, frm050203.Text1(0), "FCP") > 0 Then
        'FCP需以工程師組別列印
        strSql = strSql + " ORDER BY R01016,R01002,R01003,R01017 " '工程師組別+承辦業務+本所案號+總收文號排序
   Else
   If frm050203.Text1(9).Text = "1" Then
       '2011/9/1 modify by sonia 婧瑄說專利處案件改依智權人員部門+智權人員排序
       'strSql = strSql + " ORDER BY R01002, R01003 "
       'Modified by Lydia 2018/03/01
       'strSql = strSql + " ORDER BY S2.ST03, R01002, R01003 "
       strSql = strSql + " ORDER BY ST03, R01002, R01003 "
       pub_QL05 = pub_QL05 & ";" & Left(frm050203.Label4(2), 5) & "智權人員" 'Add By Sindy 2010/9/28
   ElseIf frm050203.Text1(9).Text = "2" Then
        'Modify by Amy 2018/02/27 Order by TO_NUMBER(REPLACE(R01010,'/','')) 會error
       'strSql = strSql + " ORDER BY TO_NUMBER(REPLACE(R01010,'/','')), R01003 "
       strSql = "Select * From (" & strSql & ") ORDER BY TO_NUMBER(REPLACE(R01010,'/','')), R01003 "
       pub_QL05 = pub_QL05 & ";" & Left(frm050203.Label4(2), 5) & "發文日"  'Add By Sindy 2010/9/28
   Else
       strSql = strSql + " ORDER BY R01001, R01003 "
       pub_QL05 = pub_QL05 & ";" & Left(frm050203.Label4(2), 5) & "承辦人"  'Add By Sindy 2010/9/28
   End If
   End If
  
   CheckOC 'Added by Lydia 2018/02/22
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
       InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/9/28
       Set MFG1.Recordset = adoRecordset
   Else
JumpEnd: 'Added by Lydia 2018/02/22
       CheckOC
       InsertQueryLog (0) 'Add By Sindy 2010/9/28
       ShowNoData
       Me.Enabled = True
       Me.Hide
       frm050203.Show
       Screen.MousePointer = vbDefault
       Exit Sub
   End If
   
   Set MFG1.Recordset = adoRecordset
   CheckOC
   With MFG1
        .TextMatrix(0, 0) = "承辦人"
        'Modify by Amy 2016/12/07
        If InStr(frm050203.Text1(0), "FCP") > 0 Or InStr(frm050203.Text1(0), "FG") > 0 Then
            .TextMatrix(0, 1) = "承辦業務"
        Else
            .TextMatrix(0, 1) = "智權人員"
        End If
        .TextMatrix(0, 2) = "本所案號"
        .TextMatrix(0, 3) = "案件名稱"
        .TextMatrix(0, 4) = "案件性質"
        .TextMatrix(0, 5) = "申請國家"
        .TextMatrix(0, 6) = "FC代/申請人國籍" 'Modify by Amy 2013/11/20 原:申請人/代理人國籍
        'Modify by Amy 2016/12/07
        .TextMatrix(0, 7) = "未請款逾月數" '原:收文日
        .ColAlignment(7) = flexAlignRightCenter
        'end 2016/12/07
        .TextMatrix(0, 8) = "發文日"
        .TextMatrix(0, 9) = "FC代理人"
        'Modify by Amy 2016/12/07
        .TextMatrix(0, 10) = ""
        .TextMatrix(0, 11) = ""
        .TextMatrix(0, 12) = ""
        .TextMatrix(0, 13) = ""
        .TextMatrix(0, 14) = "發文規費"
        .ColAlignment(14) = flexAlignRightCenter
        .TextMatrix(0, 15) = "帳單金額"
        .ColAlignment(15) = flexAlignRightCenter
        'end 2016/12/07
        .TextMatrix(0, 16) = "" 'Added by Lydia 2018/02/22
   End With
   Me.Enabled = True
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm050203a = Nothing
End Sub

'Add by Amy 2013/12/02
'傳入總收文號，取得其核稿人代號，若無抓cp14
Function Get_Ep04(ByVal strCP09 As String) As String
    Dim Rs As ADODB.Recordset
    Dim strGetEp04 As String
    Dim intR As Integer
        
    Get_Ep04 = ""
    
    strGetEp04 = "Select EP04,CP14 From CaseProgress,EngineerProgress " & _
                        "Where CP09='" & strCP09 & "' And CP09=EP02(+)"
                        
   intR = 1
   Set Rs = ClsLawReadRstMsg(intR, strGetEp04)
   If intR = 1 Then
        If IsNull(Rs.Fields("EP04")) Then
            Get_Ep04 = "" & Rs.Fields("CP14")
        Else
            Get_Ep04 = "" & Rs.Fields("EP04")
        End If
   End If
   Set Rs = Nothing
End Function

'Add by Amy 2016/12/07
Private Function GetValue(strField As String, Optional ByVal bolGrid As Boolean = False) As Integer
    Dim jj As Integer
    
    If bolGrid = True Then
        For jj = 0 To MFG1.Cols - 1
            If UCase(MFG1.TextMatrix(0, jj)) = UCase(strField) Then
                GetValue = jj
                Exit For
            End If
        Next jj
    Else
        For jj = 1 To UBound(arrField)
           If UCase(arrField(jj)) = UCase(strField) Then
              GetValue = jj
              Exit For
           End If
        Next jj
    End If
End Function

Private Sub SetPSWord()
    If bolFCPFG = True Then
        Printer.CurrentX = 300
        'Modified by Lydia 2018/02/22
        'Printer.CurrentY = 10500
        Printer.CurrentY = iMaxHeight + 300
        Printer.Print "PS ：未請款逾月數以發文日二個月為標準，不參考各案件性質設定之請款月數"
    End If
End Sub
