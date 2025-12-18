VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm210121 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權部點數分析表"
   ClientHeight    =   3190
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3190
   ScaleWidth      =   4680
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   1
      Left            =   2130
      MaxLength       =   7
      TabIndex        =   1
      Top             =   480
      Width           =   915
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Index           =   0
      Left            =   1125
      MaxLength       =   7
      TabIndex        =   0
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   1500
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2430
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3330
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   45
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grd1 
      Height          =   1245
      Left            =   30
      TabIndex        =   5
      Top             =   810
      Width           =   3105
      _ExtentX        =   5486
      _ExtentY        =   2205
      _Version        =   393216
      BackColor       =   13820671
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：4181,4201,4901科目未計入"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3210
      TabIndex        =   7
      Top             =   540
      Width           =   2300
   End
   Begin VB.Line Line1 
      X1              =   1980
      X2              =   2250
      Y1              =   615
      Y2              =   615
   End
   Begin VB.Label Label2 
      Caption         =   "統計日期："
      Height          =   180
      Left            =   90
      TabIndex        =   6
      Top             =   525
      Width           =   900
   End
End
Attribute VB_Name = "frm210121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/19 日期欄已修改
'create by nickc 2008/03/28
Option Explicit

Dim m_stdDay As String
Dim m_endDay As String

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
   If txtCloseDate(0) = "" Then
       MsgBox "日期不可以空白！", vbInformation, "操作錯誤！"
       txtCloseDate(0).SetFocus
       Exit Sub
   End If
   If txtCloseDate(1) = "" Then
       MsgBox "日期不可以空白！", vbInformation, "操作錯誤！"
       txtCloseDate(1).SetFocus
       Exit Sub
   End If
   If Val(txtCloseDate(0)) > Val(txtCloseDate(1)) Then
       MsgBox "日期範圍錯誤！", vbInformation, "操作錯誤！"
       txtCloseDate(1).SetFocus
       Exit Sub
   End If
   Screen.MousePointer = vbHourglass
   Grd1.MousePointer = flexHourglass
   StrMenu
   Grd1.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   Me.Width = mdiMain.ScaleWidth
   Me.Height = mdiMain.ScaleHeight
   MoveFormToCenter Me
   cmdExit.Left = Me.Width - 200 - cmdExit.Width
   cmdExit.Top = 30
   cmdSearch.Left = cmdExit.Left - 200 - cmdSearch.Width
   cmdSearch.Top = 30
   cmdPrint.Left = cmdSearch.Left - 200 - cmdPrint.Width
   cmdPrint.Top = 30
   Grd1.Top = 780
   Grd1.Left = 60
   Grd1.Width = Me.Width - (Grd1.Left * 4)
   Grd1.Height = Me.Height - Grd1.Top - 400
   SetGrd
   'add by sonia 2023/4/17
   If Pub_StrUserSt03 = "M51" Then
      Label1.Visible = True
   Else
      Label1.Visible = False
   End If
   'end 2023/4/17
End Sub

Private Sub SetGrd()
   'edit by nickc 2008/04/08 保留獨立，加欄位
   'Grd1.Cols = 10
   Grd1.Cols = 11
   
   Grd1.row = 0
   Grd1.col = 0: Grd1.Text = "業務區"
   Grd1.ColWidth(0) = 1000
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 1: Grd1.Text = "業務達成點數"
   Grd1.ColWidth(1) = 1300
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 2: Grd1.Text = "P"
   Grd1.ColWidth(2) = 1200
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 3: Grd1.Text = "T"
   Grd1.ColWidth(3) = 1200
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 4: Grd1.Text = "CFP"
   Grd1.ColWidth(4) = 1200
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 5: Grd1.Text = "CFT"
   Grd1.ColWidth(5) = 1200
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 6: Grd1.Text = "L"
   Grd1.ColWidth(6) = 1000
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 7: Grd1.Text = "C"
   Grd1.ColWidth(7) = 1000
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 8: Grd1.Text = "其它"
   Grd1.ColWidth(8) = 1200
   Grd1.CellAlignment = flexAlignCenterCenter
   'edit by nickc 2008/04/08 保留獨立，加欄位
   'Grd1.col = 9: Grd1.Text = ""
   'Grd1.ColWidth(9) = 0
   'Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 9: Grd1.Text = "保留"
   Grd1.ColWidth(9) = 1200
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.col = 10: Grd1.Text = ""
   Grd1.ColWidth(10) = 0
   Grd1.CellAlignment = flexAlignCenterCenter
   Grd1.ColAlignment(1) = flexAlignRightCenter
   Grd1.ColAlignment(2) = flexAlignRightCenter
   Grd1.ColAlignment(3) = flexAlignRightCenter
   Grd1.ColAlignment(4) = flexAlignRightCenter
   Grd1.ColAlignment(5) = flexAlignRightCenter
   Grd1.ColAlignment(6) = flexAlignRightCenter
   Grd1.ColAlignment(7) = flexAlignRightCenter
   Grd1.ColAlignment(8) = flexAlignRightCenter
   'add by nickc 2008/04/08 保留獨立，加欄位
   Grd1.ColAlignment(9) = flexAlignRightCenter
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210121 = Nothing
End Sub

Private Sub txtCloseDate_GotFocus(Index As Integer)
   TextInverse txtCloseDate(Index)
   CloseIme
End Sub

Private Sub txtCloseDate_KeyPress(Index As Integer, KeyAscii As Integer)
   If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 13 And KeyAscii <> 8 And KeyAscii <> 9 Then
       KeyAscii = 0
   End If
End Sub

Private Sub txtCloseDate_Validate(Index As Integer, Cancel As Boolean)
   If txtCloseDate(Index) <> "" Then
      If ChkDate(txtCloseDate(Index)) = False Then
         Cancel = True
         txtCloseDate(Index).SetFocus
         txtCloseDate_GotFocus Index
         Exit Sub
      End If
      If Index = 1 Then
        If RunNick2(txtCloseDate(0), txtCloseDate(1)) = True Then
           txtCloseDate(Index).SetFocus
           txtCloseDate_GotFocus Index
           Cancel = True
           Exit Sub
        End If
     End If
   End If
End Sub

Private Sub cmdPrint_Click()
Dim m_i As Integer
Dim m_j As Integer
Dim m_width As Long
Dim m_height As Long
Dim m_posX As Long
Dim m_posY As Long
Dim m_line As Integer
Dim m_std As Integer
Dim m_end As Integer

   Screen.MousePointer = vbHourglass
   Grd1.MousePointer = flexHourglass
   '橫印
   m_line = 22
   Printer.Orientation = 2
   Printer.Font.Size = 18
   Printer.Font.Underline = True
   Printer.FontBold = True
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("智權部點數分析表") / 2)
   Printer.CurrentY = 300
   Printer.Print "智權部點數分析表"
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("統計日期：" & ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay)) / 2)
   Printer.CurrentY = 800
   Printer.Print "統計日期：" & ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay)
   m_posY = 900
   m_height = Printer.ScaleHeight / 28
   m_width = Printer.ScaleWidth / 11
   m_posX = (m_width / 2) * -1
   m_std = 1
   m_end = Grd1.Rows - 1
   For m_j = 0 To Grd1.Cols - 1
           Printer.CurrentX = m_width * (m_j + 1) + m_posX
           Printer.CurrentY = m_height * (1) + m_posY
           Printer.Print StrToStr(Grd1.TextMatrix(0, m_j), 4)
       '直線
       Printer.Line (m_width * (m_j + 1) + m_posX, (m_height * 1) - 50 + m_posY)-(m_width * (m_j + 1) + m_posX, (m_height * (2)) - 50 + m_posY)
   Next m_j
   '橫線
   Printer.Line (m_width + m_posX, m_height * (1) - 50 + m_posY)-(m_width * (Grd1.Cols) + m_posX, m_height * (1) - 50 + m_posY)
   
   PrintBig m_std, m_end, m_line, m_height, m_width, m_posX, m_posY
   m_std = m_end + 1
   Printer.EndDoc
   Grd1.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Sub

Sub PrintBig(m_std As Integer, m_end As Integer, m_line As Integer, m_height As Long, m_width As Long, m_posX As Long, m_posY As Long)
Dim m_i As Integer
Dim m_j As Integer

   With Grd1
       For m_i = m_std To m_end
           If ((m_i - m_std + 1) Mod m_line) = 0 And m_i <> m_std And m_i <> 0 And m_std <> 0 Then
               For m_j = 0 To .Cols - 1
                   Select Case m_j
                   'edit by nickc 2008/04/08 加欄位，保留獨立
                   'Case 1, 2, 3, 4, 5, 6, 7, 8
                   Case 1, 2, 3, 4, 5, 6, 7, 8, 9
                       If m_i >= m_end - 2 Then
                       Else
                           Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80 + m_posX
                           Printer.CurrentY = m_height * ((m_line)) + m_posY
                           Printer.Print .TextMatrix(m_i, m_j)
                       End If
                   'edit by nickc 2008/04/08 加欄位，保留獨立
                   'Case 9
                   Case 10
                   Case Else
                       Printer.CurrentX = m_width * (m_j + 1) + m_posX
                       Printer.CurrentY = m_height * ((m_line)) + m_posY
                       Printer.Print .TextMatrix(m_i, m_j)
                   End Select
                   If m_i >= m_end - 2 Then
                       If m_j = 0 Then
                           '直線
                           Printer.Line (m_width * (m_j + 1) + m_posX, (m_height * (1)) - 50 + m_posY)-(m_width * (m_j + 1) + m_posX, (m_height * ((m_line - 1))) - 50 + m_posY)
                       End If
                   Else
                       '直線
                       Printer.Line (m_width * (m_j + 1) + m_posX, (m_height * (1)) - 50 + m_posY)-(m_width * (m_j + 1) + m_posX, (m_height * ((m_line) + 2)) - 50 + m_posY)
                   End If
               Next m_j
               '橫線
               Printer.Line (m_width + m_posX, (m_height * (m_line + 1)) - 50 + m_posY)-(m_width * (.Cols) + m_posX, (m_height * (m_line + 1)) - 50 + m_posY)
               Printer.Line (m_width + m_posX, (m_height * (m_line + 2)) - 50 + m_posY)-(m_width * (.Cols) + m_posX, (m_height * (m_line + 2)) - 50 + m_posY)
               '直線
               Printer.Line (m_width * (.Cols) + m_posX, (m_height * (1)) - 50 + m_posY)-(m_width * (.Cols) + m_posX, (m_height * (m_line + 2)) - 50 + m_posY)
               Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
               Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
               Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
               Printer.NewPage
               Printer.Font.Size = 18
               Printer.Font.Underline = True
               Printer.FontBold = True
               Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較") / 2)
               Printer.CurrentY = 300
               Printer.Print ChangeTStringToTDateString(m_stdDay) & "-" & ChangeTStringToTDateString(m_endDay) & " 新舊客戶收款比較"
               Printer.Font.Size = 12
               Printer.Font.Underline = False
               Printer.FontBold = False
           End If
           For m_j = 0 To .Cols - 1
               If ((m_i - m_std + 1) Mod m_line) = 0 Then
                   Printer.CurrentX = m_width * (m_j + 1)
                   'Modify By Sindy 2012/2/14
                   If (((m_i - m_std + 2) Mod m_line)) = 0 Then
                      Printer.CurrentY = m_height * m_line + m_posY
                   Else
                   '2012/2/14 End
                      Printer.CurrentY = m_height * (((m_i - m_std + 2) Mod m_line)) + m_posY
                   End If
                   Printer.Print StrToStr(.TextMatrix(((m_i - m_std + 2) Mod m_line), m_j), 4)
               Else
                   Select Case m_j
                   'edit by nickc 2008/04/08 加欄位，保留獨立
                   'Case 1, 2, 3, 4, 5, 6, 7, 8
                   Case 1, 2, 3, 4, 5, 6, 7, 8, 9
                       If m_i >= m_end - 2 Then
                       Else
                           Printer.CurrentX = m_width * (m_j + 2) - Printer.TextWidth(.TextMatrix(m_i, m_j)) - 80 + m_posX
                           'Modify By Sindy 2012/2/14
                           If (((m_i - m_std + 2) Mod m_line)) = 0 Then
                              Printer.CurrentY = m_height * m_line + m_posY
                           Else
                           '2012/2/14 End
                              Printer.CurrentY = m_height * (((m_i - m_std + 2) Mod m_line)) + m_posY
                           End If
                           Printer.Print .TextMatrix(m_i, m_j)
                       End If
                   'edit by nickc 2008/04/08 加欄位，保留獨立
                   'Case 9
                   Case 10
                   Case Else
                       Printer.CurrentX = m_width * (m_j + 1) + m_posX
                       'Modify By Sindy 2012/2/14
                       If (((m_i - m_std + 2) Mod m_line)) = 0 Then
                           Printer.CurrentY = m_height * m_line + m_posY
                       Else
                       '2012/2/14 End
                           Printer.CurrentY = m_height * (((m_i - m_std + 2) Mod m_line)) + m_posY
                       End If
                       Printer.Print .TextMatrix(m_i, m_j)
                   End Select
               End If
               If m_i >= m_end - 2 Then
                   If m_j = 0 Then
                       '直線
                       Printer.Line (m_width * (m_j + 1) + m_posX, (m_height * (1)) - 50 + m_posY)-(m_width * (m_j + 1) + m_posX, (m_height * (((m_i - m_std + 1) Mod m_line) + 2)) - 50 + m_posY)
                   End If
               Else
                   '直線
                   Printer.Line (m_width * (m_j + 1) + m_posX, (m_height * (1)) - 50 + m_posY)-(m_width * (m_j + 1) + m_posX, (m_height * (((m_i - m_std + 1) Mod m_line) + 2)) - 50 + m_posY)
               End If
           Next m_j
           '橫線
           Printer.Line (m_width + m_posX, m_height * (((m_i - m_std + 1) Mod m_line) + 1) - 50 + m_posY)-(m_width * (.Cols) + m_posX, m_height * (((m_i - m_std + 1) Mod m_line) + 1) - 50 + m_posY)
       Next m_i
       
       If m_i Mod m_line = 0 Then
           Printer.Line (m_width + m_posX, (m_height * ((m_line) + 1)) - 50 + m_posY)-(m_width * (.Cols) + m_posX, (m_height * ((m_line) + 1)) - 50 + m_posY)
           Printer.Line (m_width * (.Cols) + m_posX, (m_height * (1)) - 50 + m_posY)-(m_width * (.Cols) + m_posX, (m_height * ((m_line) + 1)) - 50 + m_posY)
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
           Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
           Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
       Else
           Printer.Line (m_width + m_posX, (m_height * (((m_i - m_std + 1) Mod m_line) + 1)) - 50 + m_posY)-(m_width * (.Cols) + m_posX, (m_height * (((m_i - m_std + 1) Mod m_line) + 1)) - 50 + m_posY)
           Printer.Line (m_width * (.Cols) + m_posX, (m_height * (1)) - 50 + m_posY)-(m_width * (.Cols) + m_posX, (m_height * (((m_i - m_std + 1) Mod m_line) + 1)) - 50 + m_posY)
           Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("第 " & Format(Printer.Page, "#0") & " 頁")
           Printer.CurrentY = (m_height * (m_line + 3)) - 50 + m_posY
           Printer.Print "第 " & Format(Printer.Page, "#0") & " 頁"
       End If
   End With
End Sub

Sub StrMenu()
Dim m_str As String
Dim m_rs As New ADODB.Recordset
Dim m_gr1 As String
Dim m_gr2 As String
Dim m_newc As Double, m_newp As Double, m_oldc As Double, m_oldp As Double, m_perc As Double, m_perp As Double
Dim m_Anewc As Double, m_Anewp As Double, m_Aoldc As Double, m_Aoldp As Double, m_Aperc As Double, m_Aperp As Double
Dim m_std As Integer
Dim m_end As Integer
Dim m_i As Integer
Dim m_j As Integer
Dim m_seekst03 As String
Dim m_seekst06 As String
Dim m_tmp As Variant
Dim m_tmp2 As Variant
Dim m_CalRow As Integer
Dim m_strALL As String   'ADD BY SONIA 2015/4/24 共用語法,不要重覆寫

   m_stdDay = txtCloseDate(0)
   m_endDay = txtCloseDate(1)
   m_str = "": m_strALL = ""
   
   'ADD BY SONIA 2015/4/24 改共用,以免每段改,但下面'比    率'要單獨改
   'edit by nickc 2008/04/08 P不含保留了
   'm_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4101',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oT,"
   'modify by sonia 2016/1/22 decode(ax205,'4131'...改decode(substr(ax205,1,4),'4131'...,decode(ax205,'4121'...改decode(substr(ax205,1,4),'4121'...
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4131',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFP,"
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFT,"
   '2015/4/24 4161投資法務收入也計入oL
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4141',decode(ax207, 0, ax206 * -1, ax207),'4161',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oL,"
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4151',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oC,"
   '2015/4/24 4171,4172 FCP收入及FCT收入也計入oOther
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4171', decode(ax207, 0, ax206 * -1, ax207),'4172', decode(ax207, 0, ax206 * -1, ax207),'7121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oOther,"
   'add by nickc 2008/04/08 保留獨立
   'modify by sonia 2015/4/24 加4194
   m_strALL = m_strALL & " to_char(sum(decode(substr(ax205,1,4),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),'4194',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oKeep,"
   '2015/4/24 END
   
   '抓資料 北區、中區
   m_str = m_str & " select a0902,"
   
   m_str = m_str & " to_char(sum(decode(ax207, 0, ax206 * -1, ax207)),'999,999,999,990.99') as oAll,"
   
'MODIFY BY SONIA 2015/4/24 改用共用
'   'edit by nickc 2008/04/08 P不含保留了
'   'm_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4101',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oT,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4131',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFP,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFT,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4141',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oL,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4151',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oC,"
'   m_str = m_str & " to_char(sum(decode(ax205,'7121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oOther,"
'   'add by nickc 2008/04/08 保留獨立
'   'modify by sonia 2015/4/24 加4194
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),'4194',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oKeep,"
   m_str = m_str & m_strALL
'2015/4/24 END

   m_str = m_str & " st15 As oSort"
   m_str = m_str & " From acc021, acc020, staff, customer, acc090"
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   m_str = m_str & " where ax201 = a0201 and ax202 = a0202 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205='7121')"
   m_str = m_str & " and a0205 >= '" & m_stdDay & "' and a0205 <= '" & m_endDay & "' and st01(+)=ax209 and cu01(+)=substr(ax208,1,8) and st15=a0901(+)"
   'modify by sonia 2021/1/14 +F4104~F4107
   m_str = m_str & " and cu02(+)=substr(ax208,9,1) and ax209 not in ('F4102','F4103','F4101','F4104','F4105','F4106','F4107') and substr(st15,1,2) in ('S1','S2')"
   m_str = m_str & " group by a0902,st15"
   
   
   '台北所合計、台中所合計、台南所、高雄所、其它合計
   m_str = m_str & " union select decode(substr(st15,1,2),'S1','台北所合計','S2','台中所合計','S3','台南所','S4','高雄所','其它合計') as oGRP,"
   m_str = m_str & " to_char(sum(decode(ax207, 0, ax206 * -1, ax207)),'999,999,999,990.99') as oAll,"
   
'MODIFY BY SONIA 2015/4/24 改用共用
'   'edit by nickc 2008/04/08 P不含保留了
'   'm_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4101',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oT,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4131',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFP,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFT,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4141',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oL,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4151',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oC,"
'   m_str = m_str & " to_char(sum(decode(ax205,'7121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oOther,"
'   'add by nickc 2008/04/08 保留獨立
'   'modify by sonia 2015/4/24 加4194
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),'4194',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oKeep,"
   m_str = m_str & m_strALL
'2015/4/24 END
   
   m_str = m_str & " decode(substr(st15,1,2),'S1','S1Z','S2','S2Z','S3','S3Z','S4','S4Z','SVZ') as oSort"
   m_str = m_str & " From acc021, acc020, staff, customer"
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   m_str = m_str & " where ax201 = a0201 and ax202 = a0202 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205='7121')"
   m_str = m_str & " and a0205 >= '" & m_stdDay & "' and a0205 <= '" & m_endDay & "' and st01(+)=ax209 and cu01(+)=substr(ax208,1,8)"
   'modify by sonia 2021/1/14 +F4104~F4107
   m_str = m_str & " and cu02(+)=substr(ax208,9,1) and ax209 not in ('F4102', 'F4103', 'F4101','F4104','F4105','F4106','F4107')"
   m_str = m_str & " group by decode(substr(st15,1,2),'S1','台北所合計','S2','台中所合計','S3','台南所','S4','高雄所','其它合計'),decode(substr(st15,1,2),'S1','S1Z','S2','S2Z','S3','S3Z','S4','S4Z','SVZ') "
   
   '國內合計
   m_str = m_str & " union select '國內合計' as oGRP,"
   m_str = m_str & " to_char(sum(decode(ax207, 0, ax206 * -1, ax207)),'999,999,999,990.99') as oAll,"
   
'MODIFY BY SONIA 2015/4/24 改用共用
'   'edit by nickc 2008/04/08 P不含保留了
'   'm_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4101',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oT,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4131',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFP,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFT,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4141',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oL,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4151',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oC,"
'   m_str = m_str & " to_char(sum(decode(ax205,'7121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oOther,"
'   'add by nickc 2008/04/08 保留獨立
'   'modify by sonia 2015/4/24 加4194
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),'4194',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oKeep,"
   m_str = m_str & m_strALL
'2015/4/24 END
   
   m_str = m_str & " 'SWZ' as oSort"
   m_str = m_str & " From acc021, acc020, staff, customer"
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   m_str = m_str & " where ax201 = a0201 and ax202 = a0202 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205='7121')"
   m_str = m_str & " and a0205 >= '" & m_stdDay & "' and a0205 <= '" & m_endDay & "' and st01(+)=ax209 and cu01(+)=substr(ax208,1,8)"
   'modify by sonia 2021/1/14 +F4104~F4107
   m_str = m_str & " and cu02(+)=substr(ax208,9,1) and ax209 not in ('F4102', 'F4103', 'F4101','F4104','F4105','F4106','F4107') "
   '比    率
   'edit by nickc 2008/04/11
   'm_str = m_str & " union select oGRP ,'  ',to_char(decode(oAll,0,0,op/oall*100),'90.99')||'%',to_char(decode(oAll,0,0,ot/oall*100),'90.99')||'%'"
   m_str = m_str & " union select oGRP ,to_char(oAll,'999,999,999,990.99'),to_char(decode(oAll,0,0,op/oall*100),'90.99')||'%',to_char(decode(oAll,0,0,ot/oall*100),'90.99')||'%'"
   
   m_str = m_str & " ,to_char(decode(oAll,0,0,ocfp/oall*100),'90.99')||'%',to_char(decode(oAll,0,0,ocft/oall*100),'90.99')||'%'"
   m_str = m_str & " ,to_char(decode(oAll,0,0,ol/oall*100),'90.99')||'%',to_char(decode(oAll,0,0,oc/oall*100),'90.99')||'%'"
   'edit by nickc 2008/04/08 保留獨立
   'm_str = m_str & " ,to_char(decode(oAll,0,0,oother/oall*100),'90.99')||'%',osort from ("
   m_str = m_str & " ,to_char(decode(oAll,0,0,oother/oall*100),'90.99')||'%'"
   'edit by nickc 2008/04/11
   'm_str = m_str & " ,to_char(decode(oAll,0,0,oKeep/oall*100),'90.99')||'%',osort from ("
   m_str = m_str & " ,' ',osort from ("
   
   m_str = m_str & " select '比    率' as oGRP,"
   'edit by nickc 2008/04/11 不要含保留
   'm_str = m_str & " to_char(sum(decode(ax207, 0, ax206 * -1, ax207)),'9999999990.99') as oAll,"
   'modify by sonia 2015/4/24 加4194
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4191',0,'4192',0,'4194',0,decode(ax207, 0, ax206 * -1, ax207))),'999999999990.99') as oAll,"
   
   'edit by nickc 2008/04/08 P不含保留了
   'm_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),0)),'9999999990.99') as oP,"
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oP,"
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4101',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oT,"
   'modify by sonia 2016/1/22 decode(ax205,'4131'...改decode(substr(ax205,1,4),'4131'...,decode(ax205,'4121'...改decode(substr(ax205,1,4),'4121'...
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4131',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oCFP,"
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4121',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oCFT,"
   '2015/4/24 4161投資法務收入也計入oL
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4141',decode(ax207, 0, ax206 * -1, ax207),'4161',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oL,"
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4151',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oC,"
   '2015/4/24 4171,4172 FCP收入及FCT收入也計入oOther
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4171', decode(ax207, 0, ax206 * -1, ax207),'4172', decode(ax207, 0, ax206 * -1, ax207),'7121',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oOther,"
   'add by nickc 2008/04/08 保留獨立
   'modify by sonia 2015/4/24 加4194
   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),'4194',decode(ax207, 0, ax206 * -1, ax207),0)),'999999999990.99') as oKeep,"
   
   m_str = m_str & " 'SXZ' as oSort"
   m_str = m_str & " From acc021, acc020, staff, customer"
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   m_str = m_str & " where ax201 = a0201 and ax202 = a0202 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205='7121')"
   m_str = m_str & " and a0205 >='" & m_stdDay & "'  and a0205 <='" & m_endDay & "' and st01(+)=ax209 and cu01(+)=substr(ax208,1,8)"
   'modify by sonia 2021/1/14 +F4104~F4107
   m_str = m_str & " and cu02(+)=substr(ax208,9,1) and ax209 not in ('F4102', 'F4103', 'F4101','F4104','F4105','F4106','F4107')"
   m_str = m_str & " ) tmpA"
   '加空白
   'add by nickc 2008/04/08 保留獨立
   'm_str = m_str & " union select ' ','','','','','','','','','SYZ' as oSort from dual "
   m_str = m_str & " union select ' ','','','','','','','','','','SYZ' as oSort from dual "
   '全所合計
   m_str = m_str & " union select '全所合計' as oGRP,"
   m_str = m_str & " to_char(sum(decode(ax207, 0, ax206 * -1, ax207)),'999,999,999,990.99') as oAll,"
   
'MODIFY BY SONIA 2015/4/24 改用共用
'   'edit by nickc 2008/04/08 P不含保留了
'   'm_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4111',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oP,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4101',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oT,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4131',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFP,"
'   m_str = m_str & " to_char(sum(decode(ax205,'4121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oCFT,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4141',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oL,"
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4151',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oC,"
'   m_str = m_str & " to_char(sum(decode(ax205,'7121',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oOther,"
'   'add by nickc 2008/04/08 保留獨立
'   'modify by sonia 2015/4/24 加4194
'   m_str = m_str & " to_char(sum(decode(substr(ax205,1,4),'4191',decode(ax207, 0, ax206 * -1, ax207),'4192',decode(ax207, 0, ax206 * -1, ax207),'4194',decode(ax207, 0, ax206 * -1, ax207),0)),'999,999,999,990.99') as oKeep,"
   m_str = m_str & m_strALL
'2015/4/24 END
   
   m_str = m_str & " 'SZZ' as oSort"
   m_str = m_str & " From acc021, acc020, staff, customer"
   'Modify by Amy 2019/08/01 增加創新業務組用收入 420101 原:substr(ax205, 1, 2) = '41'
   m_str = m_str & " where ax201 = a0201 and ax202 = a0202 and ax209 is not null and (substr(ax205, 1, 1) = '4' Or ax205='7121')"
   m_str = m_str & " and a0205 >= '" & m_stdDay & "' and a0205 <= '" & m_endDay & "' and st01(+)=ax209 and cu01(+)=substr(ax208,1,8)"
   m_str = m_str & " and cu02(+)=substr(ax208,9,1) "
   
   m_str = m_str & "order by oSort "
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
       Grd1.Visible = False
       Set Grd1.Recordset = m_rs
       SetGrd
   '    With m_rs
   
   ''統計及計算
       With Grd1
           '台北所合計
           For m_i = m_std To .Rows - 1
               If .TextMatrix(m_i, 0) = "台北所合計" Then
                   m_end = m_i
                   m_seekst06 = m_seekst06 & m_end & ","
                   Exit For
               End If
           Next m_i
           For m_i = 0 To 9
               .row = m_end
               .col = m_i
               .CellBackColor = &HFF80FF   '&HFF00FF
           Next m_i
           '台中所合計
           For m_i = m_std To .Rows - 1
               If .TextMatrix(m_i, 0) = "台中所合計" Then
                   m_end = m_i
                   m_seekst06 = m_seekst06 & m_end & ","
                   Exit For
               End If
           Next m_i
           For m_i = 0 To 9
               .row = m_end
               .col = m_i
               .CellBackColor = &HFF80FF   '&HFF00FF
           Next m_i
           '台南所
           For m_i = m_std To .Rows - 1
               If .TextMatrix(m_i, 0) = "台南所" Then
                   m_end = m_i
                   m_seekst06 = m_seekst06 & m_end & ","
                   Exit For
               End If
           Next m_i
           For m_i = 0 To 9
               .row = m_end
               .col = m_i
               .CellBackColor = &HFF80FF   '&HFF00FF
           Next m_i
           '高雄所
           For m_i = m_std To .Rows - 1
               If .TextMatrix(m_i, 0) = "高雄所" Then
                   m_end = m_i
                   m_seekst06 = m_seekst06 & m_end & ","
                   Exit For
               End If
           Next m_i
           For m_i = 0 To 9
               .row = m_end
               .col = m_i
               .CellBackColor = &HFF80FF   '&HFF00FF
           Next m_i
           '國內合計
           For m_i = m_std To .Rows - 1
               If .TextMatrix(m_i, 0) = "國內合計" Then
                   m_end = m_i
                   m_CalRow = m_i
                   Exit For
               End If
           Next m_i
           For m_i = 0 To 9
               .row = m_end
               .col = m_i
               .CellBackColor = &HFFFF80
           Next m_i
           '全所合計
           For m_i = m_std To .Rows - 1
               If .TextMatrix(m_i, 0) = "全所合計" Then
                   m_end = m_i
                   Exit For
               End If
           Next m_i
   '        .TextMatrix(m_end, 2) = Val(.TextMatrix(m_end - 2, 2)) + Val(.TextMatrix(m_end - 1, 2))
   '        .TextMatrix(m_end, 3) = Val(.TextMatrix(m_end - 2, 3)) + Val(.TextMatrix(m_end - 1, 3))
   '        .TextMatrix(m_end, 5) = Val(.TextMatrix(m_end - 2, 5)) + Val(.TextMatrix(m_end - 1, 5))
   '        .TextMatrix(m_end, 6) = Val(.TextMatrix(m_end - 2, 6)) + Val(.TextMatrix(m_end - 1, 6))
   '        .TextMatrix(m_end, 8) = Val(.TextMatrix(m_end - 2, 8)) + Val(.TextMatrix(m_end - 1, 8))
   '        .TextMatrix(m_end, 9) = Val(.TextMatrix(m_end - 2, 9)) + Val(.TextMatrix(m_end - 1, 9))
   '        .TextMatrix(m_end, 4) = Format((Val(.TextMatrix(m_end, 3)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
   '        .TextMatrix(m_end, 7) = Format((Val(.TextMatrix(m_end, 6)) / Val(.TextMatrix(m_end, 9))) * 100, "0.00") & "%"
           For m_i = 0 To 9
               .row = m_end
               .col = m_i
               .CellBackColor = &H80FF80
           Next m_i
           .Rows = .Rows + 2
           .TextMatrix(.Rows - 1, 0) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 1) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 2) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 3) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 4) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 5) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 6) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 7) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 8) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           'add by nickc 2008/04/08 加欄位，保留獨立
           .TextMatrix(.Rows - 1, 9) = Replace(Replace(StrConv("P_" & Trim(.TextMatrix(m_CalRow, 2)) & ":T_" & Trim(.TextMatrix(m_CalRow, 3)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 3)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 2), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 3), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           
           .MergeCells = flexMergeRestrictRows
           .MergeRow(.Rows - 1) = True
           .MergeCol(0) = True
           .MergeCol(1) = True
           .MergeCol(2) = True
           .MergeCol(3) = True
           .MergeCol(4) = True
           .MergeCol(5) = True
           .MergeCol(6) = True
           .MergeCol(7) = True
           .MergeCol(8) = True
           .Rows = .Rows + 1
           .TextMatrix(.Rows - 1, 0) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 1) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 2) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 3) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 4) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 5) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 6) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 7) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           .TextMatrix(.Rows - 1, 8) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           'add by nickc 2008/04/08 加欄位，保留獨立
           .TextMatrix(.Rows - 1, 9) = Replace(Replace(StrConv("CFP_" & Trim(.TextMatrix(m_CalRow, 4)) & ":CFT_" & Trim(.TextMatrix(m_CalRow, 5)) & "=" & IIf(Val(.TextMatrix(m_CalRow, 5)) = 0, "0", Format(Val(Replace(.TextMatrix(m_CalRow, 4), ",", "")) / Val(Replace(.TextMatrix(m_CalRow, 5), ",", "")), "###,###,###,##0.00")) & ":1", vbWide), StrConv(".", vbWide), "."), StrConv(",", vbWide), ",")
           
           .MergeCells = flexMergeRestrictRows
           .MergeRow(.Rows - 1) = True
           .MergeCol(0) = True
           .MergeCol(1) = True
           .MergeCol(2) = True
           .MergeCol(3) = True
           .MergeCol(4) = True
           .MergeCol(5) = True
           .MergeCol(6) = True
           .MergeCol(7) = True
           .MergeCol(8) = True
           'add by nickc 2008/04/08 加欄位，保留獨立
           .MergeCol(9) = True
           
       End With
       Grd1.Visible = True
   Else
       ShowNoData
       txtCloseDate(0).SetFocus
   End If
End Sub
