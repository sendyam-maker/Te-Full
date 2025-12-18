VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm170206 
   BorderStyle     =   1  '單線固定
   Caption         =   "同仁婚喪互助明細表"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5475
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   5475
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   4440
      TabIndex        =   7
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3420
      TabIndex        =   6
      Top             =   90
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   3
      Top             =   4110
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label21 
         Caption         =   "印表機"
         Height          =   315
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox Textwf01 
      Height          =   270
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   0
      Top             =   450
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3135
      Left            =   30
      TabIndex        =   1
      Top             =   885
      Width           =   5400
      _ExtentX        =   9525
      _ExtentY        =   5530
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "婚喪互助日期："
      Height          =   180
      Left            =   30
      TabIndex        =   2
      Top             =   480
      Width           =   1260
   End
End
Attribute VB_Name = "frm170206"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2008/12/23 add by sonia
'2008/12/30 add by Sindy
Option Explicit
Dim i As Integer, j As Integer
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
'Dim PaperX As Double
'Dim paperY As Double
Dim iPgae As Integer, iLine As Integer
Dim strType As String


Private Sub cmdok_Click(Index As Integer)
Dim strEmpID As String
   
   Select Case Index
      Case 0 '列印
          If Textwf01 = "" Then
             MsgBox "婚喪互助日期不可以空白！", vbInformation, "操作錯誤！"
             Textwf01.SetFocus
             Exit Sub
         End If
         If Textwf01 <> "" Then
             If ChkDate(Textwf01) = False Then
                Textwf01.SetFocus
                Exit Sub
             End If
'             If ChkWork(ChangeTStringToWString(Textwf01)) = False Then
'                Textwf01.SetFocus
'                Exit Sub
'             End If
         End If
         
         strEmpID = ""
         For i = 1 To GRD1.Rows - 1
            GRD1.col = 0
            GRD1.row = i
            If Trim(GRD1.Text) = "V" Then
               GRD1.col = 2
               If strEmpID = "" Then
                  strEmpID = "'" & GRD1.Text & "'"
               Else
                  strEmpID = strEmpID & ",'" & GRD1.Text & "'"
               End If
            End If
         Next i
         If strEmpID <> "" Then
            Screen.MousePointer = vbHourglass
            m_StrSQL = " AND wfa01=" & ChangeTStringToWString(Textwf01.Text) & " AND wfa02 in (" & strEmpID & ")"
            Call StrMenu1
            Screen.MousePointer = vbDefault
         Else
            MsgBox "未點選任何婚喪互助名單！", vbInformation
            Exit Sub
         End If
         
      Case 1 '結束
         Unload Me
   End Select
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
   Set frm170206 = Nothing
End Sub

Private Sub grd1_SelChange()
Dim i As Integer
   
   GRD1.Visible = False
   GRD1.row = GRD1.MouseRow
   
   GRD1.col = 0
   If GRD1.row <> 0 Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
   GRD1.Visible = True
   cmdok(0).Enabled = True
   cmdok(0).Default = True
End Sub

Private Sub textWF01_GotFocus()
   InverseTextBox Textwf01
   cmdok(0).Enabled = False
End Sub

Private Sub textWF01_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textWF01_Validate(Cancel As Boolean)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   If Textwf01 <> "" Then
      'Modified by Morgan 2022/10/31 +wf11
      strSql = "SELECT '',sqldateT(WF01),wf02,st02,wf03||' '||decode(wf03,'1','婚','2','喪','')||decode(wf11,'1','(父親)','2','(母親)','3','(配偶)','4','(兒子)','5','(女兒)',''),sqldateT(WF04) FROM WeddingAndFuneral,staff " & _
               "where WF02=st01(+) and WF01=" & DBDATE(Textwf01) & " order by WF01,WF02 "
      If rsTmp.State = 1 Then rsTmp.Close
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount = 0 Then
         MsgBox "無此日期的婚喪互助名單！", vbInformation
         Cancel = True
         textWF01_GotFocus
      Else
         cmdok(0).Enabled = True
         cmdok(0).Default = True
      End If
      Set GRD1.Recordset = rsTmp
      SetGrd
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Sub

Private Sub SetGrd()
Dim arrGridHeadText, arrGridHeadWidth
Dim iRow As Integer
   
   arrGridHeadText = Array("V", "日期", "互助同仁", "姓名", "原因", "扣款日期")
   arrGridHeadWidth = Array(200, 800, 1000, 1000, 1000, 1000)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub


Sub StrMenu1()
Dim dblAmt As Double

Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 1 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

m_str = "SELECT wfa02,a.ST02,wfa03,b.ST02,wfa04,wf03 " & _
                "FROM WFAmount,Staff a,Staff b,WeddingAndFuneral " & _
              "WHERE wfa05='1' " & _
                    "AND wfa02=a.ST01(+) " & _
                    "AND wfa03=b.ST01(+) " & _
                    "AND wf01=wfa01 AND wf02=wfa02" & m_StrSQL & " Order BY wfa02,wfa03 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        m_rs.MoveFirst
        
        iLine = 1
        'PrintTitle '列印表頭
        strType = "" '切頁條件
        dblAmt = 0
        Do While Not m_rs.EOF
            
            For m_i = 1 To 6
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("wfa02"))
            strTemp(2) = CheckStr(m_rs.Fields(1))
            strTemp(3) = CheckStr(m_rs.Fields("wfa03"))
            strTemp(4) = CheckStr(m_rs.Fields(3))
            strTemp(5) = CheckStr(m_rs.Fields("wfa04"))
            strTemp(6) = CheckStr(m_rs.Fields("wf03"))
            If strTemp(6) = "1" Then strTemp(6) = "婚"
            'Modified by Morgan 2022/10/31
            'If strTemp(6) = "2" Then strTemp(6) = "喪"
            If strTemp(6) = "2" Then strTemp(6) = Mid(GRD1.TextMatrix(GRD1.row, 4), 3)
            'end 2022/10/31
            
            If iLine = 1 Then PrintTitle   '列印表頭
            
            If strType <> "" Then
               If iLine > 50 Or (strType <> strTemp(1)) Then
                  
                   If (strType <> strTemp(1)) Then
                      Printer.CurrentX = 500
                      Printer.CurrentY = iLine * 300
                      Printer.Print String(140, "-")
                      
                      iLine = iLine + 1
                      Printer.CurrentX = PLeft(2)
                      Printer.CurrentY = iLine * 300
                      Printer.Print "合　計："
                      Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt, "##,##0"))
                      Printer.CurrentY = iLine * 300
                      Printer.Print Format(dblAmt, "##,##0")
                      
                      dblAmt = 0 '合　計
                   End If
                   
                   'If .AbsolutePosition <> .RecordCount Then
                       Printer.NewPage
                       iLine = 1
                       PrintTitle '列印表頭
                   'End If
               End If
            End If
            
            PrintDetail '列印表中
            
            strType = strTemp(1)
            dblAmt = dblAmt + strTemp(5)
            m_rs.MoveNext
        Loop
         
         '列印表尾
         Printer.CurrentX = 500
         Printer.CurrentY = iLine * 300
         Printer.Print String(140, "-")
         
         iLine = iLine + 1
         Printer.CurrentX = PLeft(2)
         Printer.CurrentY = iLine * 300
         Printer.Print "合　計："
         Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(dblAmt, "##,##0"))
         Printer.CurrentY = iLine * 300
         Printer.Print Format(dblAmt, "##,##0")
    End With
Else
    MsgBox "無符合列印的資料!!!", vbExclamation + vbOKOnly
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintTitle()
GetPleft

'PaperX = 12000
'paperY = 7500

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("同仁婚喪互助明細表") / 2)
Printer.CurrentY = iLine * 300
Printer.Print "同仁婚喪互助明細表"

iLine = iLine + 1
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "列印人：" & strUserName
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
Printer.CurrentY = iLine * 300
Printer.Print "頁　　次：" & Printer.Page

iLine = iLine + 2
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print "同仁：" & strTemp(2)
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("互助原因：" & strTemp(6)) / 2)
Printer.CurrentY = iLine * 300
Printer.Print "互助原因：" & strTemp(6)

iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "員工代號"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　　名"
Printer.CurrentX = PLeft(3) - Printer.TextWidth("互助金額")
Printer.CurrentY = iLine * 300
Printer.Print "互助金額"

iLine = iLine + 1
Printer.CurrentX = 500
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")

iLine = iLine + 1
End Sub

Sub GetPleft()
PLeft(1) = 3500
PLeft(2) = 5500
PLeft(3) = 8500
End Sub

Sub PrintDetail()
   '員工代號
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   '姓　　名
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(4)
   '互助金額
   Printer.CurrentX = PLeft(3) - Printer.TextWidth(Format(strTemp(6), "##,##0"))
   Printer.CurrentY = iLine * 300
   Printer.Print Format(strTemp(5), "##,##0")
   
   iLine = iLine + 1
End Sub
