VERSION 5.00
Begin VB.Form frm170224 
   BorderStyle     =   1  '單線固定
   Caption         =   "例外扣繳項目員工名單"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4770
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4770
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   0
      Top             =   2370
      Width           =   4665
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   1
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   2
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2610
      TabIndex        =   4
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3630
      TabIndex        =   5
      Top             =   60
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "包含下列情形之員工名單：  　　　　　　　　　　　　1. 有輸入所得稅率者"
      Height          =   720
      Index           =   0
      Left            =   1005
      TabIndex        =   3
      Top             =   930
      Width           =   2200
   End
End
Attribute VB_Name = "frm170224"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Memo by Morgan 2010/12/2 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
'2009/2/5 add by sonia
Option Explicit
Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         StrMenu
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Sub StrMenu()

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1 '1.直印 2.橫印
   'Modified by Morgan 2013/1/21 取消是否適用一般勞健保費率(原sd11 改為 勞健保是否以合夥人身分投保)
   'm_str = "SELECT SD01,ST02,SD08,SD11 FROM Staff,SalaryData " & _
            "WHERE SD01<'F' AND (SD08 IS NOT NULL OR SD11='N') " & _
            "AND SD01=st01(+) ORDER BY SD01"
   m_str = "SELECT SD01,ST02,SD08 FROM Staff,SalaryData " & _
            "WHERE SD01<'F' AND SD08 IS NOT NULL " & _
            "AND SD01=st01(+) ORDER BY SD01"
   If m_rs.State = 1 Then m_rs.Close
   m_rs.CursorLocation = adUseClient
   m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
   If Not m_rs.EOF And Not m_rs.BOF Then
      With m_rs
         m_rs.MoveFirst
         
         iLine = 1
         PrintTitle '列印表頭
         strType = "" '切頁條件
         Do While Not m_rs.EOF
             
            For m_i = 1 To 10
               strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields("SD01"))
            strTemp(2) = CheckStr(m_rs.Fields("ST02"))
            strTemp(3) = CheckStr(m_rs.Fields("SD08")) '所得稅率
            'Removed by Morgan 2013/1/21
            'strTemp(4) = CheckStr(m_rs.Fields("SD11")) '是否適用一般勞健保費率
            
            PrintDetail '列印表中
            
            m_rs.MoveNext
         Loop
          
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
   
   Printer.Font.Size = 12
   Printer.Font.Underline = False
   Printer.FontBold = False
   
   Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("例外扣繳項目員工名單") / 2)
   Printer.CurrentY = iLine * 300
   Printer.Print "例外扣繳項目員工名單"
   
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
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * 300
   Printer.Print "員工代號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print "姓　名"
   Printer.CurrentX = PLeft(3) - Printer.TextWidth("所得稅率")
   Printer.CurrentY = iLine * 300
   Printer.Print "所得稅率"
   
'Removed by Morgan 2013/1/21
'   Printer.CurrentX = PLeft(4) - Printer.TextWidth("是否適用一般勞健保費率")
'   Printer.CurrentY = iLine * 300
'   Printer.Print "是否適用一般勞健保費率"
   
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print String(140, "-")
   
   iLine = iLine + 1
End Sub

Sub GetPleft()
   PLeft(1) = 2500
   PLeft(2) = 4000
   PLeft(3) = 6000
   PLeft(4) = 9000
End Sub

Sub PrintDetail()
   Printer.CurrentX = PLeft(1) + 300
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(1)
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(2)
   Printer.CurrentX = PLeft(3) - 500
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(3)
   Printer.CurrentX = PLeft(4) - 1500
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(4)
   
   iLine = iLine + 1
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSQL As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSQL = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSQL Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170224 = Nothing
End Sub
