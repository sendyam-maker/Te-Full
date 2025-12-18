VERSION 5.00
Begin VB.Form frm083015 
   BorderStyle     =   1  '單線固定
   Caption         =   "開拓客戶地址條"
   ClientHeight    =   2580
   ClientLeft      =   2310
   ClientTop       =   2100
   ClientWidth     =   4380
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4380
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   228
      TabIndex        =   8
      Top             =   1548
      Width           =   3795
      Begin VB.ComboBox Combo1 
         Height          =   276
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   9
         Top             =   180
         Width           =   2880
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.OptionButton Opt1 
      Caption         =   " 國       籍："
      Height          =   255
      Index           =   1
      Left            =   228
      TabIndex        =   3
      Top             =   1188
      Width           =   1215
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "屬性代號："
      Height          =   255
      Index           =   0
      Left            =   228
      TabIndex        =   0
      Top             =   708
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2520
      Style           =   1  '圖片外觀
      TabIndex        =   6
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3336
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1428
      MaxLength       =   3
      TabIndex        =   1
      Top             =   708
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   2868
      MaxLength       =   3
      TabIndex        =   2
      Top             =   708
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1428
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1188
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   2868
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1188
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   2628
      X2              =   2748
      Y1              =   828
      Y2              =   828
   End
   Begin VB.Line Line2 
      X1              =   2628
      X2              =   2748
      Y1              =   1308
      Y2              =   1308
   End
End
Attribute VB_Name = "frm083015"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/3 日期欄已修改
Option Explicit

Dim SeekPrint As Integer
Dim SeekPrintL As Integer
Dim m_count As Integer
'Add By Cheng 2002/09/26
Dim blnClkSure As Boolean '判斷是否按下確定按鈕

Private Sub cmdBack_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   'Add By Cheng 2002/09/26
   blnClkSure = False
   If Opt1(0).Value = True Then
      'Modify By Cheng 2002/09/10
      If ChkRange(Text1(0), Text1(1), "屬性代號") = False Then
         blnClkSure = True
         Me.Text1(0).SetFocus
         Text1_GotFocus 0
         Exit Sub
      End If
   Else
      'Modify By Cheng 2002/09/10
      If ChkRange(Text1(2), Text1(3), "國籍") = False Then
         blnClkSure = True
         Me.Text1(2).SetFocus
         Text1_GotFocus 2
         Exit Sub
      End If
   End If
   
  If Combo1.ListIndex >= SeekPrint Then
     m_count = Combo1.ListIndex + 1
  Else
     m_count = Combo1.ListIndex
  End If
  Set Printer = Printers(m_count)
  Printer.Orientation = 1
  DoEvents
  PrintCase
  If Err.Number = 0 Then
     MsgBox "列印完成!", vbInformation, "開拓客戶地址條"
  End If
End Sub

Private Sub PrintCase()
 Dim i As Integer, St As String, iPrint As Integer
 Dim IntF As Integer
On Error GoTo ErrHand
   If Opt1(0).Value = True Then
      If Text1(0) = "" And Text1(1) <> "" Then
         St = " AND ECD02 <='" & Text1(1) & "'"
      ElseIf Text1(0) <> "" And Text1(1) <> "" Then
         St = " AND (ECD02 BETWEEN '" & Text1(0) & "' AND '" & Text1(1) & "')"
      End If
   Else
      If Text1(2) = "" And Text1(3) <> "" Then
         St = " AND ECD09 <='" & Text1(3) & "'"
      ElseIf Text1(2) <> "" And Text1(3) <> "" Then
         St = " AND (ECD09 BETWEEN '" & Text1(2) & "' AND '" & Text1(3) & "')"
      End If
   End If
   strExc(0) = "SELECT ECD04,ECD05,ECD06,ECD07,ECD08,DECODE(ECD09,NA01,NA04)," & _
      "SUBSTR(ECD10,1,30),SUBSTR(ECD10,31,30),SUBSTR(ECD03,1,30)," & _
      "SUBSTR(ECD03,31,30),ECD02||ECD01 FROM EXPANDCUSDETAIL,NATION WHERE " & _
      "ECD09=NA01(+) " & St & " ORDER BY ECD02,ECD01"
   If RsTemp.State = adStateOpen Then RsTemp.Close
   RsTemp.Open strExc(0), cnnConnection
   If RsTemp.EOF And RsTemp.BOF Then
      MsgBox "資料庫內無資料 !", vbInformation
      Exit Sub
   End If
   
'   Printer.FontSize = 8
   Printer.FontSize = 12
   Printer.Height = 3000
   Printer.Width = 10000
   iPrint = 1
   With RsTemp
      Do While Not .EOF
         For i = 0 To 10
            Printer.CurrentX = 1000
'            Printer.CurrentY = i * 160
            Printer.CurrentY = i * 220
            If IsNull(.Fields(i)) = False Then Printer.Print .Fields(i)
         Next
         Printer.CurrentX = 4200
'         Printer.CurrentY = (i - 1) * 160
         Printer.CurrentY = (i - 1) * 220
         Printer.Print Format(iPrint, "000000")
         iPrint = iPrint + 1
         Printer.NewPage
         .MoveNext
      Loop
   End With
   Printer.EndDoc
   Exit Sub
ErrHand:
   MsgBox Err.Description
End Sub

Private Sub Form_Load()
   Dim strPrint As String
   Dim i As Integer

   m_count = 0

   MoveFormToCenter Me
   Opt1_Click 0
   
'*****************
'印表設定
'*****************
   strPrint = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
       Set Printer = Printers(i)
       If Printer.DeviceName <> strPrint Then
           Combo1.AddItem Printer.DeviceName, m_count
           m_count = m_count + 1
       End If
       If Printer.DeviceName = strPrint Then
           SeekPrint = i
       End If
   Next i
   Combo1.Text = Combo1.List(0)
   
End Sub

Private Sub Opt1_Click(Index As Integer)
 Dim i As Integer
On Error Resume Next
   For i = 0 To 3
      Text1(i).Enabled = False
   Next
   Select Case Index
      Case 0
         Text1(0).Enabled = True
         Text1(1).Enabled = True
         Text1(0).SetFocus
      Case 1
         Text1(2).Enabled = True
         Text1(3).Enabled = True
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
   'Add By Cheng 2002/09/26
   Select Case Index
   Case 1 '屬性代號
      If blnClkSure = False Then
         If Me.Text1(0).Text <> "" And Me.Text1(1).Text <> "" Then
            If Me.Text1(0).Text > Me.Text1(1).Text Then
               MsgBox "屬性代號範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Text1(0).SetFocus
               Text1_GotFocus 0
               Exit Sub
            End If
         End If
      Else
         blnClkSure = False
      End If
   Case 3 '國籍
      If blnClkSure = False Then
         If Me.Text1(2).Text <> "" And Me.Text1(3).Text <> "" Then
            If Me.Text1(2).Text > Me.Text1(3).Text Then
               MsgBox "國籍範圍輸入錯誤!!!", vbExclamation + vbOKOnly
               Me.Text1(2).SetFocus
               Text1_GotFocus 2
               Exit Sub
            End If
         End If
      Else
         blnClkSure = False
      End If
   End Select
   
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
 Dim strTempName As String, St As String
   If Text1(Index) = "" Then Exit Sub
   Select Case Index
      Case 0, 1
'         If objLawDll.GetExpand(Text1(Index), strTempName) = False Then Cancel = True
      Case 2, 3
'         If objPublicData.GetNation(Text1(Index), strTempName) = False Then Cancel = True
   End Select
   If Cancel Then TextInverse Text1(Index)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set Printer = Printers(SeekPrint)
   Printer.Orientation = SeekPrintL
   Set frm083015 = Nothing
End Sub
