VERSION 5.00
Begin VB.Form frm040328 
   BorderStyle     =   1  '單線固定
   Caption         =   "核准領證期限表"
   ClientHeight    =   915
   ClientLeft      =   8145
   ClientTop       =   1470
   ClientWidth     =   4185
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   915
   ScaleWidth      =   4185
   Begin VB.TextBox txtDate 
      Height          =   264
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   240
      Width           =   945
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   2325
      TabIndex        =   1
      Top             =   195
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   350
      Index           =   1
      Left            =   3150
      TabIndex        =   2
      Top             =   195
      Width           =   800
   End
   Begin VB.Label Label1 
      Caption         =   "收文日："
      Height          =   180
      Index           =   1
      Left            =   315
      TabIndex        =   3
      Top             =   300
      Width           =   765
   End
End
Attribute VB_Name = "frm040328"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
Dim PLeft(0 To 12) As Integer, iPrint As Integer, iPage As Integer, strTemp(1 To 5) As String
Dim m_iTitleFontSize As Single, m_iFontSize As Single
Dim m_iStartX As Integer, m_iStartY As Integer
Dim m_iPageHeight As Integer, m_iLineHeight As Integer, m_stTmp As String
Dim m_iMargin As Integer, m_Title As String

Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/2 清除查詢印表記錄檔欄位
         Process
         Screen.MousePointer = vbDefault
      Case 1 '結束
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
  
   MoveFormToCenter Me
   txtDate = strSrvDate(2)
   m_Title = Me.Caption
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm040328 = Nothing
End Sub

Private Sub txtDate_GotFocus()
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtDate.IMEMode = 2
   CloseIme
   TextInverse txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If ChkDate(txtDate) = False Then
      Cancel = True
   End If
End Sub

Private Sub Process()
   
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txtDate 'Add By Sindy 2010/12/2
   strSql = "select NP02||'-'||NP03||'-'||NP04||'-'||NP05 本所案號, substrB(PA11,1,16) 申請案號" & _
      ", '領證及繳年費' 下一程序,CP05-19110000 收文日, NP08-19110000 本所期限" & _
      " From caseprogress, nextprogress, patent" & _
      " where CP01='P' and CP05=" & TransDate(txtDate, 2) & " and cp14='" & strUserNum & "'" & _
      " and np01(+)=CP09 AND NP07='601'" & _
      " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04" & _
      " order by 1"
   
On Error GoTo ErrHnd

   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/2
         If DoPrint(AdoRecordSet3) = True Then
            MsgBox "列印完畢！", vbInformation
         End If
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/2
         MsgBox "無待列印資料！", vbExclamation
      End If
   End With
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Sub

Private Function DoPrint(ByRef p_Recordset As ADODB.Recordset) As Boolean

   Dim iCol As Integer, iRecs As Integer
   
On Error GoTo ErrHnd

   GetPleft
   iRecs = 0
   iPage = 1
   With p_Recordset
      .MoveFirst
      Do While Not .EOF
         iRecs = iRecs + 1
         If iRecs = 1 Then
            PrintPageHeader
            PrintPageHeader1
         End If
         For iCol = LBound(strTemp) To UBound(strTemp)
            Select Case iCol
               Case 4, 5
                  strTemp(iCol) = Format("" & .Fields(iCol - 1), "###/##/##")
               Case Else
                  strTemp(iCol) = "" & .Fields(iCol - 1)
            End Select
         Next
         
         PrintDetail
         .MoveNext
      Loop
      If iRecs > 0 Then
         Call PrintReportFooter(iRecs)
      End If
   End With
   DoPrint = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

Sub GetPleft()

   Printer.Orientation = 1
   m_iTitleFontSize = 22
   m_iFontSize = 12
   m_iStartX = 500
   m_iStartY = 500
   m_iPageHeight = Printer.ScaleHeight
   m_iLineHeight = 300
   m_iMargin = 500
   
   Erase PLeft
   PLeft(0) = 500
   '本所案號(2000)
   PLeft(1) = 500
   '申請案號(1500)
   PLeft(2) = PLeft(1) + 2000
   '下一程序(1900)
   PLeft(3) = PLeft(2) + 1500
   '收文日(1500)
   PLeft(4) = PLeft(3) + 1900
   '本所期限(1500)
   PLeft(5) = PLeft(4) + 1500
    
End Sub

Sub PrintPageHeader()

   iPrint = m_iStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = m_iTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = (Printer.ScaleWidth - Printer.TextWidth(m_Title)) / 2
   Printer.CurrentY = iPrint
   Printer.Print m_Title
   iPrint = iPrint + 500
   Printer.Font.Size = m_iFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   PrintNewLine
   
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   PrintNewLine
   
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "收文日期：" & txtDate
   
   Printer.CurrentX = Printer.ScaleWidth - m_iMargin - 2500
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   PrintNewLine
End Sub

Sub PrintPageHeader1()

   Call PrintNewLine(False, 1)
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "下一程序"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "收文日"
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "本所期限"
   
   PrintNewLine
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print String(95, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

   Call PrintNewLine(True, 3)
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print String(95, "-")
   PrintNewLine
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = iPrint
   Printer.Print "合計： " & iRecCount & " 筆"
   PrintMemo
   Printer.EndDoc
   
End Sub

Sub PrintDetail()

   Dim iCol As Integer
   
   PrintNewLine
   For iCol = LBound(strTemp) To UBound(strTemp)
      Printer.CurrentX = PLeft(iCol)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(iCol)
   Next
    
End Sub

Private Sub PrintMemo()
   Printer.Font.Size = 10
   Printer.CurrentX = m_iStartX
   Printer.CurrentY = m_iPageHeight - m_iMargin - Printer.TextHeight("註")
   Printer.Print ""
   Printer.Font.Size = m_iFontSize
End Sub

Private Sub PrintNewLine(Optional ByVal p_bolHeader1 As Boolean = True, Optional ByVal p_iExtraLines As Integer = 2)

   iPrint = iPrint + m_iLineHeight
   If iPrint >= (m_iPageHeight - m_iMargin - p_iExtraLines * m_iLineHeight) Then
      Printer.CurrentX = m_iStartX
      Printer.CurrentY = iPrint
      Printer.Print String(132, "-")
      PrintMemo
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If p_bolHeader1 Then
         PrintPageHeader1
      End If
      iPrint = iPrint + m_iLineHeight
   End If
   
End Sub
