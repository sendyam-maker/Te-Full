VERSION 5.00
Begin VB.Form frm160110 
   BorderStyle     =   1  '單線固定
   Caption         =   "端午、中秋獎金名單"
   ClientHeight    =   3024
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4968
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3024
   ScaleWidth      =   4968
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   8
      Top             =   2400
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   2
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   9
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   1
      Top             =   1260
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   2040
      MaxLength       =   5
      TabIndex        =   0
      Top             =   930
      Width           =   555
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3915
      TabIndex        =   4
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2970
      TabIndex        =   3
      Top             =   60
      Width           =   915
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "(1.端午  2.中秋)"
      Height          =   180
      Left            =   2370
      TabIndex        =   7
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "獎金類別："
      Height          =   180
      Left            =   1110
      TabIndex        =   6
      Top             =   1320
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "獎金年月："
      Height          =   180
      Left            =   1110
      TabIndex        =   5
      Top             =   960
      Width           =   900
   End
End
Attribute VB_Name = "frm160110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
Option Explicit

Dim m_StrSQL As String
Dim m_str  As String
Dim m_str2  As String
Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_i As Integer
Dim PLeft(1 To 7) As Integer
Dim strTemp(1 To 7) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
Dim dblCnt As Double, dblTotCnt As Double
Dim dblAmt As Double, dblTotAmt As Double

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
   Case 0
           If txt1(0) = "" Then
               MsgBox "獎金年月不可以空白！", vbCritical, "操作錯誤！"
               txt1(0).SetFocus
               Exit Sub
           End If
           If txt1(1) = "" Then
               MsgBox "獎金類別不可以空白！", vbInformation, "操作錯誤！"
               txt1(1).SetFocus
               Exit Sub
           End If
           
           Screen.MousePointer = vbHourglass
           m_StrSQL = ""
           If txt1(0) <> "" Then
               m_StrSQL = m_StrSQL & " and ob01='" & Mid(DBDATE(txt1(0) & "01"), 1, 6) & "' "
           End If
           If txt1(1) <> "" Then
               m_StrSQL = m_StrSQL & " and ob02='" & txt1(1) & "' "
           End If
           StrMenu
           Screen.MousePointer = vbDefault
   Case 1
           Unload Me
   Case Else
   End Select
   Printer.Font.Size = 12
End Sub

Sub StrMenu()
   'modify by sonia 2018/9/14 劉經理要求台一投資R04單獨印
   m_str = "select ob03,st02,ob04,ob05,sqldatet(st13),a0802,a0801 " & _
                "from staff,ohBonus,SalaryData,acc080 " & _
                "where ob03=st01(+) and st03<>'R04' " & _
                "and ob03=sd01(+) " & _
                "and sd19=a0801(+) and ob05<>0 " & m_StrSQL & _
                "order by a0801 asc,ob04 desc,ob03 ASC "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       PrintData
   Else
       ShowNoData
       'Exit Sub
   End If
   'add by sonia 2018/9/14 劉經理要求台一投資R04單獨印
   m_str = "select ob03,st02,ob04,ob05,sqldatet(st13),a0802,a0801 " & _
                "from staff,ohBonus,SalaryData,acc080 " & _
                "where ob03=st01(+) and st03='R04' " & _
                "and ob03=sd01(+) " & _
                "and sd19=a0801(+) and ob05<>0 " & m_StrSQL & _
                "order by a0801 asc,ob04 desc,ob03 ASC "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       PrintData
   Else
       MsgBox "資料庫中臺一投資搜尋不到符合資料!!", , "沒有資料"   'modify by sonia 2021/2/25 改名稱
       Exit Sub
   End If
   ShowPrintOk
   'end 2018/9/14
End Sub

Sub PrintData()
Dim int_i As Integer
Dim dblYear As Double, dblMonth As Double
Dim intRow As Integer
Dim BolChkEnd04 As Boolean 'Add By Sindy 2014/12/29 最後一筆是否為04留職停薪

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1
With m_rs
    .MoveFirst
    
    strType = ""
    'PrintTitle
    intRow = 0
    dblCnt = 0
    dblTotCnt = 0
    dblAmt = 0
    dblTotAmt = 0
    
    Do While Not .EOF
        
        For m_i = 1 To 7
            strTemp(m_i) = ""
        Next m_i
        
        strTemp(1) = CheckStr(.Fields(0))
        strTemp(2) = CheckStr(.Fields(1))
        strTemp(6) = CheckStr(.Fields(5)) '公司名稱
        strTemp(7) = CheckStr(.Fields(6)) '公司別
        
         '為固定任職期間寫死的條件判斷
         If strTemp(1) = "63001" Or strTemp(1) = "63002" Or _
            strTemp(1) = "64001" Or strTemp(1) = "65001" Or _
            strTemp(1) = "68001" Or strTemp(1) = "72010" Then
            If strTemp(1) = "63001" Then strTemp(3) = "63/02/22 -- 65/09/30"
            If strTemp(1) = "63002" Then strTemp(3) = "63/02/22 -- 65/09/30"
            If strTemp(1) = "64001" Then strTemp(3) = "64/03/01 -- 65/09/30"
            If strTemp(1) = "65001" Then strTemp(3) = "65/05/01 -- 65/09/30"
            If strTemp(1) = "68001" Then strTemp(3) = "65/04/01 -- 65/09/30"
            If strTemp(1) = "72010" Then strTemp(3) = "72/04/20 -- 77/06/30"
            If intRow >= 15 Or _
               (strType <> strTemp(6)) Then
                 'If .AbsolutePosition <> .RecordCount Then   2009/2/5 CANCEL BY SONIA
                     If strType <> strTemp(6) And strType <> "" Then
                        Call PrintEnd
                        dblCnt = 0
                        dblAmt = 0
                     End If
                     If strType <> "" Then Printer.NewPage: intRow = 0
                     iLine = 1
                     Call PrintTitle
                 'End If
            End If
            PrintDetail
            strType = strTemp(6)
         End If
         
         strTemp(4) = PUB_ChangeNianZi(Val(CheckStr(.Fields(2))))
         strTemp(5) = Format(CheckStr(.Fields(3)), "###,###,###")
         strTemp(3) = CheckStr(.Fields(4)) '到職日
        
         'Add By Sindy 2014/12/29 檢查最後一筆是否為04留職停薪
         BolChkEnd04 = False
         strSql = "select * from staff_change where sc01='" & strTemp(1) & "' " & _
                  "and sc02=(select max(sc02) from staff_change where sc01='" & strTemp(1) & "') " & _
                  "and sc03='04'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            BolChkEnd04 = True
         End If
         '2014/12/29 END
        
         '任職時間
         m_str2 = "select sqldatet(sc02) as 日期,sc03 " & _
                        "from staff_change " & _
                        "where sc03 in ('02','03','04','08','09','10')  " & _
                        "and sc01='" & strTemp(1) & "' " & _
                        "order by sc02 "
         If m_rs2.State = 1 Then m_rs2.Close
         m_rs2.CursorLocation = adUseClient
         m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs2.EOF And Not m_rs2.BOF Then
            m_rs2.MoveFirst
            int_i = 0
            Do While Not m_rs2.EOF
               int_i = int_i + 1
               
               '有到職日並且異動資料只有一筆
               If strTemp(3) <> "" And m_rs2.RecordCount = 1 Then
                   If Val(ChangeTDateStringToTString(strTemp(3))) <= Val(ChangeWStringToTString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "01"))), "yyyy") & "1231")) Then
                       If CheckStr(m_rs2.Fields(1)) = "04" Then
                            strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                      Else
                            strTemp(3) = strTemp(3) & " -- " & ChangeWStringToTDateString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "01"))), "yyyy") & "1231")
                      End If
                   Else
                       strTemp(3) = strTemp(3) & " -- "
                   End If
                   If intRow >= 15 Or _
                    (strType <> strTemp(6)) Then
                        'If .AbsolutePosition <> .RecordCount Then   2009/2/5 CANCEL BY SONIA
                            If strType <> strTemp(6) And strType <> "" Then
                               Call PrintEnd
                               dblCnt = 0
                               dblAmt = 0
                            End If
                            If strType <> "" Then Printer.NewPage: intRow = 0
                            iLine = 1
                            PrintTitle
                        'End If
                   End If
                   PrintDetail
                   strType = strTemp(6)
                   
               '有到職日並且非異動資料的最後一筆
               ElseIf strTemp(3) <> "" And m_rs2.RecordCount > 1 And m_rs2.AbsolutePosition <> m_rs2.RecordCount Then
                   If m_rs2.AbsolutePosition <> 1 Then
                       strTemp(1) = ""
                       strTemp(2) = ""
                   End If
                   strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                   strTemp(4) = ""
                   strTemp(5) = ""
                   If int_i Mod 2 <> 0 Then
                         If intRow >= 15 Or _
                          (strType <> strTemp(6)) Then
                              'If .AbsolutePosition <> .RecordCount Then   2009/2/5 CANCEL BY SONIA
                                  If strType <> strTemp(6) And strType <> "" Then
                                     Call PrintEnd
                                     dblCnt = 0
                                     dblAmt = 0
                                  End If
                                  If strType <> "" Then Printer.NewPage: intRow = 0
                                  iLine = 1
                                  PrintTitle
                              'End If
                         End If
                         PrintDetail
                         strType = strTemp(6)
                   End If
                   strTemp(3) = PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0)))
                   
               Else
                   strTemp(1) = ""
                   strTemp(2) = ""
                   strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                   If int_i Mod 2 <> 0 Then
                         
                     'Add By Sindy 2014/12/29
                     If BolChkEnd04 = True Then
                        strTemp(4) = PUB_ChangeNianZi(Val(CheckStr(.Fields(2))))
                        strTemp(5) = Format(CheckStr(.Fields(3)), "###,###,###")
                     End If
                     '2014/12/29 END
                         
                     If intRow >= 15 Or _
                      (strType <> strTemp(6)) Then
                          'If .AbsolutePosition <> .RecordCount Then   2009/2/5 CANCEL BY SONIA
                              If strType <> strTemp(6) And strType <> "" Then
                                 Call PrintEnd
                                 dblCnt = 0
                                 dblAmt = 0
                              End If
                              If strType <> "" Then Printer.NewPage: intRow = 0
                              iLine = 1
                              PrintTitle
                          'End If
                     End If
                     PrintDetail
                         
                     'Add By Sindy 2014/12/29
                     If BolChkEnd04 = True Then
                        Exit Do
                     End If
                     '2014/12/29 END
                     
                     strType = strTemp(6)
                   End If
                   strTemp(3) = PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0)))
                   
                   If Val(ChangeTDateStringToTString(strTemp(3))) <= Val(ChangeWStringToTString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "01"))), "yyyy") & "1231")) Then
                       If CheckStr(m_rs2.Fields(1)) = "04" Then
                            strTemp(3) = strTemp(3) & " -- " & ChangeTStringToTDateString(ChangeWDateStringToTString(DateAdd("d", -1, ChangeTStringToWDateString(ChangeTDateStringToTString(PUB_ScDateWriteDeal(CheckStr(.Fields(0)), CheckStr(m_rs2.Fields(0))))))))
                      Else
                            strTemp(3) = strTemp(3) & " -- " & ChangeWStringToTDateString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "01"))), "yyyy") & "1231")
                      End If
                   Else
                       strTemp(3) = strTemp(3) & " -- "
                   End If
                   strTemp(4) = PUB_ChangeNianZi(Val(CheckStr(.Fields(2))))
                   strTemp(5) = Format(CheckStr(.Fields(3)), "###,###,###")
                   If intRow >= 15 Or _
                    (strType <> strTemp(6)) Then
                        'If .AbsolutePosition <> .RecordCount Then   2009/2/5 CANCEL BY SONIA
                            If strType <> strTemp(6) And strType <> "" Then
                               Call PrintEnd
                               dblCnt = 0
                               dblAmt = 0
                            End If
                            If strType <> "" Then Printer.NewPage: intRow = 0
                            iLine = 1
                            PrintTitle
                        'End If
                   End If
                   PrintDetail
                   strType = strTemp(6)
               End If
               m_rs2.MoveNext
            Loop
         '沒異動資料
         Else
            '為固定任職期間寫死的條件判斷
            If strTemp(1) = "63001" Or strTemp(1) = "63002" Or _
              strTemp(1) = "64001" Or strTemp(1) = "65001" Or _
              strTemp(1) = "68001" Or strTemp(1) = "72010" Then
              If strTemp(1) = "72010" Then strTemp(3) = "77/07/01"
              strTemp(1) = "": strTemp(2) = ""
            End If
            
             If Val(ChangeTDateStringToTString(strTemp(3))) <= Val(ChangeWStringToTString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "01"))), "yyyy") & "1231")) Then
                 strTemp(3) = strTemp(3) & " -- " & ChangeWStringToTDateString(Format(DateAdd("yyyy", -1, ChangeWStringToWDateString(DBDATE(txt1(0) & "01"))), "yyyy") & "1231")
             Else
                 strTemp(3) = strTemp(3) & " -- "
             End If
             If intRow >= 15 Or _
                      (strType <> strTemp(6)) Then
                 'If .AbsolutePosition <> .RecordCount Then   2009/2/5 CANCEL BY SONIA否則最後一筆不會跳頁
                     If strType <> strTemp(6) And strType <> "" Then
                        Call PrintEnd
                        dblCnt = 0
                        dblAmt = 0
                     End If
                     If strType <> "" Then Printer.NewPage: intRow = 0
                     iLine = 1
                     PrintTitle
                 'End If
             End If
             PrintDetail
             strType = strTemp(6)
         End If
         
         Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B
         iLine = iLine + 1
         intRow = intRow + 1
         dblCnt = dblCnt + 1           '人數小計
         dblTotCnt = dblTotCnt + 1 '人數合計
         dblAmt = dblAmt + CheckStr(.Fields(3))           '金額小計
         dblTotAmt = dblTotAmt + CheckStr(.Fields(3)) '金額合計
        .MoveNext
    Loop
    Call PrintEnd
    Printer.CurrentX = 5500
    Printer.CurrentY = iLine * 300
    Printer.Print "合　計"
    Printer.CurrentX = PLeft(4) - Printer.TextWidth(dblTotCnt & " 人")
    Printer.CurrentY = iLine * 300
    Printer.Print dblTotCnt & " 人"
    Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblTotAmt, "###,###,###"))
    Printer.CurrentY = iLine * 300
    Printer.Print Format(dblTotAmt, "###,###,###")
    iLine = iLine + 1
    Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B
End With
Printer.EndDoc
'ShowPrintOk
End Sub

Sub PrintEnd()
Printer.CurrentX = 5500
Printer.CurrentY = iLine * 300
Printer.Print "小　計"
Printer.CurrentX = PLeft(4) - Printer.TextWidth(dblCnt & " 人")
Printer.CurrentY = iLine * 300
Printer.Print dblCnt & " 人"
Printer.CurrentX = PLeft(5) - Printer.TextWidth(Format(dblAmt, "###,###,###"))
Printer.CurrentY = iLine * 300
Printer.Print Format(dblAmt, "###,###,###")
iLine = iLine + 1
Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B
iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 1000
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 8500
PLeft(5) = 10000
End Sub

Sub PrintTitle()
Dim oStr As String
oStr = Mid(DBDATE(txt1(0) & "01"), 1, 4) - 1911 & "年度" & IIf(Val(txt1(1)) = 1, "端午", "中秋") & "禮金明細表"

GetPleft

Printer.Font.Size = 16
Printer.Font.Underline = False
Printer.FontBold = False

iLine = iLine + 2
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(oStr) / 2)
Printer.CurrentY = iLine * 300
Printer.Print oStr

iLine = iLine + 2
Printer.Font.Size = 12
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.Font.Size = 14
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "公司別：" & strTemp(7) & " " & strTemp(6)
Printer.Font.Size = 12
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 3
Printer.Font.Size = 14
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "員工編號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "任　職　時　間"
Printer.CurrentX = PLeft(4) - Printer.TextWidth("年　資")
Printer.CurrentY = iLine * 300
Printer.Print "年　資"
Printer.CurrentX = PLeft(5) - Printer.TextWidth("獎金金額")
Printer.CurrentY = iLine * 300
Printer.Print "獎金金額"
iLine = iLine + 1
'Printer.CurrentX = 500
'Printer.CurrentY = iLine * 300
'Printer.Print String(140, "-")
Printer.Line (500, iLine * 300)-(11000, iLine * 300), , B
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 5
   If m_j = 4 Or m_j = 5 Then
      Printer.CurrentX = PLeft(m_j) - Printer.TextWidth(strTemp(m_j))
   Else
      Printer.CurrentX = PLeft(m_j)
   End If
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j)
Next m_j
iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm160110 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
   Case 0
           If txt1(Index) <> "" Then
               If ChkDate(txt1(Index) & "01") = False Then
                   Cancel = True
                   Exit Sub
               End If
           End If
   Case 1
           If txt1(Index) <> "" Then
               Select Case txt1(Index)
               Case "1", "2"
               Case Else
                   MsgBox "獎金類別只可以輸入 1 或 2！", vbInformation, "輸入錯誤！"
                   Cancel = True
                   Exit Sub
               End Select
           End If
   End Select
End Sub
