VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm040208 
   BorderStyle     =   1  '單線固定
   Caption         =   "CF 結餘單案件明細查詢"
   ClientHeight    =   1580
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1580
   ScaleWidth      =   4380
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   4725
      Left            =   90
      TabIndex        =   10
      Top             =   1620
      Width           =   8805
      _ExtentX        =   15522
      _ExtentY        =   8326
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   3
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   3
      Top             =   1200
      Width           =   195
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   2
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   2
      Top             =   870
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   1
      Left            =   1080
      MaxLength       =   7
      TabIndex        =   1
      Top             =   870
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   285
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   540
      Width           =   3195
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   3270
      TabIndex        =   5
      Top             =   60
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   2160
      TabIndex        =   4
      Top             =   60
      Width           =   1005
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.查詢 2.列印)"
      Height          =   180
      Left            =   1350
      TabIndex        =   9
      Top             =   1245
      Width           =   1155
   End
   Begin VB.Line Line1 
      X1              =   1590
      X2              =   2370
      Y1              =   990
      Y2              =   990
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "輸出方式："
      Height          =   180
      Left            =   60
      TabIndex        =   8
      Top             =   1252
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "結餘日期："
      Height          =   180
      Left            =   60
      TabIndex        =   7
      Top             =   922
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   90
      TabIndex        =   6
      Top             =   592
      Width           =   900
   End
End
Attribute VB_Name = "frm040208"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/05 Form2.0已修改 Printer列印未改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/7/26 日期欄已修改
Option Explicit

Dim m_StrSQL As String
Dim tmpRs As ADODB.Recordset
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
   Select Case Index
   Case 0
        If txt1(3) = "" Then
            MsgBox "輸出方式不可以空白！", vbInformation, "操作錯誤！"
            txt1(3).SetFocus
            Exit Sub
        End If
        Cancel = False
        txt1_Validate 0, Cancel
        If Cancel = True Then Exit Sub
        txt1_Validate 1, Cancel
        If Cancel = True Then Exit Sub
        txt1_Validate 2, Cancel
        If Cancel = True Then Exit Sub
        txt1_Validate 3, Cancel
        If Cancel = True Then Exit Sub
        
        ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/21 清除查詢印表記錄檔欄位
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & " and a240005 in (" & GetAddStr(txt1(0)) & ") "
            pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/12/21
        Else
            m_StrSQL = m_StrSQL & " and a240005 in (" & GetAddStr(Systemkind_g) & ") "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and a240001 >='" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and a240001 <='" & txt1(2) & "' "
        End If
        If txt1(1) <> "" Or txt1(2) <> "" Then
           pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/21
        End If
        If txt1(3).Text = "1" Then
           pub_QL05 = pub_QL05 & ";" & Label3 & "1.查詢" 'Add By Sindy 2010/12/21
        ElseIf txt1(3).Text = "2" Then
           pub_QL05 = pub_QL05 & ";" & Label3 & "2.列印" 'Add By Sindy 2010/12/21
        End If
        
        Screen.MousePointer = vbHourglass
        StrMenu
        Screen.MousePointer = vbDefault
        If tmpRs.RecordCount <> 0 Then
            Select Case txt1(3)
            Case "1"
                    Me.Hide
                    frm040208_1.Show
                    Set frm040208_1.grd1.Recordset = tmpRs
                    SetGrd frm040208_1.grd1
            Case "2"
                    PrintData
            End Select
        End If
        If tmpRs.State = 1 Then tmpRs.Close
   Case 1
        Unload Me
   End Select
End Sub

Sub StrMenu()
Dim m_rs As New ADODB.Recordset
Dim m_str As String

   m_str = "select sqldatet(a240001) 結餘日,a240002 結餘單號,a240005||'-'||a240006||'-'||a240007||'-'||a240008 本所案號,a240011 申請國家,nvl(decode(pa09,'020',p1.ptm04,p1.ptm03),decode(tm10,'020',p2.ptm04,p2.ptm03)) 專利種類,to_char(nvl(a241006,0),'999,999,999') 浮動金,to_char(nvl(a241007,0),'999,999,999') 結餘金額,a240012 代理人,a240009 申請人,st02 智權人員 from acc240,acc241,patent,patenttrademarkmap p1,trademark,patenttrademarkmap p2,staff where a240002=a241001(+) and '998'=a241002(+) and a240005=pa01(+) and a240006=pa02(+) and a240007=pa03(+) and a240008=pa04(+) and a240005=tm01(+) and a240006=tm02(+) and a240007=tm03(+) and a240008=tm04(+) and '1'=p1.ptm01(+) and pa08=p1.ptm02(+) and '2'=p2.ptm01(+) and tm08=p2.ptm02(+) and a240010=st01(+) " & m_StrSQL & " order by a240005,a240006,a240007,a240008 "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       InsertQueryLog (m_rs.RecordCount) 'Add By Sindy 2010/12/21
       Set grd1.Recordset = m_rs
       Set tmpRs = m_rs
       SetGrd grd1
   Else
       InsertQueryLog (0) 'Add By Sindy 2010/12/21
       ShowNoData
       Set tmpRs = m_rs
       SetGrd grd1
   End If
End Sub

Sub PrintData()
Dim m_i As Long
Dim m_j As Integer

   If grd1.Rows = 1 And grd1.TextMatrix(1, 1) = "" Then Exit Sub
   Printer.Orientation = 2
   PrintTitle
   For m_i = 1 To grd1.Rows - 1
       For m_j = 0 To grd1.Cols - 1
           strTemp(m_j + 1) = grd1.TextMatrix(m_i, m_j)
           Select Case m_j
           Case 3, 4
                   strTemp(m_j + 1) = StrToStr(strTemp(m_j + 1), 4)
           Case 7, 8
                   strTemp(m_j + 1) = StrToStr(strTemp(m_j + 1), 15)
           End Select
       Next m_j
       PrintDetail
       If iLine >= 35 Then
           Printer.NewPage
           PrintTitle
       End If
   Next m_i
   Printer.EndDoc
   ShowPrintOk
End Sub

Sub SetGrd(oGrd As MSHFlexGrid)
   oGrd.ColWidth(0) = 800
   oGrd.ColWidth(1) = 1050
   oGrd.ColWidth(2) = 1400
   oGrd.ColWidth(3) = 900
   oGrd.ColWidth(4) = 900
   oGrd.ColAlignment(5) = flexAlignRightCenter
   oGrd.ColAlignment(6) = flexAlignRightCenter
   oGrd.ColWidth(7) = 3000
   oGrd.ColWidth(8) = 2500
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(0) = Systemkind_g
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040208 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Dim strTemp As Variant
Dim strTemp1 As Variant
Dim i, s, j

   Select Case Index
      Case 0
        If Len(Trim(txt1(0))) <> 0 Then
               strTemp = Split(Systemkind_g, ",")
               strTemp1 = Split(txt1(0), ",")
               For i = 0 To UBound(strTemp1)
                   s = 0
                   For j = 0 To UBound(strTemp)
                       If strTemp1(i) = strTemp(j) Then
                           s = 1
                       End If
                   Next j
                   If s = 0 Then
                       s = MsgBox(strUserNum + " 沒有 " + strTemp1(i) + " 的使用權限 ", , "USER 權限不足!!!")
                       txt1(0).SetFocus
                       txt1_GotFocus (0)
                       Cancel = True
                       Exit Sub
                   End If
               Next i
           End If
      Case 1, 2
         If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
            Me.txt1(Index).SetFocus
            txt1_GotFocus Index
            Cancel = True
            Exit Sub
         End If
         If Index = 2 Then
              If Not nickChgRan(txt1(1), txt1(2), "准駁日") Then
                 txt1(1).SetFocus
                 txt1_GotFocus 1
                 Exit Sub
              End If
          End If
      Case 3
           Select Case txt1(3)
           Case "", "1", "2"
           Case Else
                   MsgBox "輸出方式只可以輸入 1 或 2 ！", vbInformation, "操作錯誤！"
                   txt1(3).SetFocus
                   txt1_GotFocus 3
                   Cancel = True
                   Exit Sub
           End Select
      Case Else
   End Select
End Sub

Sub PrintTitle()
   GetPleft
   Printer.Font.Size = 18
   Printer.Font.Underline = True
   Printer.FontBold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("CF 結餘案件明細查詢") / 2)
   Printer.CurrentY = 300
   Printer.Print "CF 結餘案件明細查詢"
   Printer.Font.Size = 10
   Printer.Font.Underline = False
   Printer.FontBold = False
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = 900
   Printer.Print "系  統  別：" & IIf(txt1(0) = "", Systemkind_g, txt1(0))
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
   Printer.CurrentY = 900
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = 1200
   Printer.Print "結餘日期：" & ChangeTStringToTDateString(txt1(1)) & "-" & ChangeTStringToTDateString(txt1(2))
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
   Printer.CurrentY = 1200
   Printer.Print "頁　　次：" & Printer.Page
   iLine = 6
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "結餘日期"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "結餘單號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print "申請國家"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iLine * 300
   Printer.Print "種　　類"
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iLine * 300
   Printer.Print "浮  動  金"
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iLine * 300
   Printer.Print "結餘金額"
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iLine * 300
   Printer.Print "代理人"
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iLine * 300
   Printer.Print "申請人"
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iLine * 300
   Printer.Print "智權人員"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print String(8, "=")
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print String(10, "=")
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * 300
   Printer.Print String(13, "=")
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * 300
   Printer.Print String(8, "=")
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iLine * 300
   Printer.Print String(8, "=")
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iLine * 300
   Printer.Print String(8, "=")
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iLine * 300
   Printer.Print String(8, "=")
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iLine * 300
   Printer.Print String(31, "=")
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iLine * 300
   Printer.Print String(31, "=")
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iLine * 300
   Printer.Print String(10, "=")
   iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer

   For m_j = 1 To 10
       If m_j = 6 Or m_j = 7 Then
           Printer.CurrentX = PLeft(m_j + 1) - 200 - Printer.TextWidth(strTemp(m_j))
       Else
           Printer.CurrentX = PLeft(m_j)
       End If
       Printer.CurrentY = iLine * 300
       Printer.Print strTemp(m_j)
   Next m_j
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 200
   PLeft(2) = 1200
   PLeft(3) = 2400
   PLeft(4) = 3900
   PLeft(5) = 4900
   PLeft(6) = 5900
   PLeft(7) = 6900
   PLeft(8) = 8000
   PLeft(9) = 11500
   PLeft(10) = 15000
End Sub
