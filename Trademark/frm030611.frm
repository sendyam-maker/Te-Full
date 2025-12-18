VERSION 5.00
Begin VB.Form frm030611 
   BorderStyle     =   1  '單線固定
   Caption         =   "表一～表五"
   ClientHeight    =   1650
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4170
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   4170
   Begin VB.TextBox textTMBM07_1 
      Height          =   264
      Left            =   1320
      MaxLength       =   5
      TabIndex        =   0
      Top             =   720
      Width           =   1092
   End
   Begin VB.TextBox textTMBM07_2 
      Height          =   264
      Left            =   2880
      MaxLength       =   5
      TabIndex        =   1
      Top             =   720
      Width           =   1092
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   2220
      TabIndex        =   3
      Top             =   60
      Width           =   912
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   3180
      TabIndex        =   4
      Top             =   60
      Width           =   912
   End
   Begin VB.ComboBox cmbPrinter 
      Height          =   300
      ItemData        =   "frm030611.frx":0000
      Left            =   1320
      List            =   "frm030611.frx":0002
      TabIndex        =   2
      Top             =   1710
      Visible         =   0   'False
      Width           =   2652
   End
   Begin VB.Label Label1 
      Caption         =   "公報卷期："
      Height          =   252
      Left            =   240
      TabIndex        =   6
      Top             =   720
      Width           =   972
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   2760
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label10 
      Caption         =   "印表機 :"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   1710
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frm030611"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo By Sindy 2010/8/11 日期欄已修改
Option Explicit

'edit by nick 2004/12/14
'Dim m_DefaultPrinter As String

Private Sub Form_Load()
'   Dim Prn As Printer
'edit by nick 2004/12/14
'   m_DefaultPrinter = Printer.DeviceName
   MoveFormToCenter Me
'edit by nick 2004/12/14
'   For Each Prn In Printers
'      If Prn.DeviceName <> m_DefaultPrinter Then
'         cmbPrinter.AddItem Prn.DeviceName
'      End If
'   Next
'   cmbPrinter.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'edit by nick 2004/12/14
'   Dim Prn As Printer
'   For Each Prn In Printers
'      If Prn.DeviceName = m_DefaultPrinter Then
'         Set Printer = Prn
'         Exit For
'      End If
'   Next
   'Add By Cheng 2002/07/19
   Set frm030611 = Nothing
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdok_Click()
   Dim strFrom As String
   Dim strTo As String
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   If CheckDataValid() = True Then
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/22 清除查詢印表記錄檔欄位
      If Len(textTMBM07_1) <> 0 Or Len(textTMBM07_2) <> 0 Then
         pub_QL05 = pub_QL05 & ";" & Label1 & textTMBM07_1 & "-" & textTMBM07_2 'Add By Sindy 2010/10/22
      End If
      
      strFrom = textTMBM07_1
      strTo = textTMBM07_2
      frm030606.PrintReportBK cmbPrinter.Text, strFrom, strTo
      frm030608.PrintReportBK cmbPrinter.Text, strFrom, strTo
      frm030609.PrintReportBK cmbPrinter.Text, strFrom, strTo
      frm030610.PrintReportBK cmbPrinter.Text, strFrom, strTo
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      
      InsertQueryLog ("") 'Add By Sindy 2010/10/22
      
      strTit = "輸出報表"
      strMsg = "列印結束"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
End Sub

' 公報卷期(起)
Private Sub textTMBM07_1_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(textTMBM07_1) = False Then
      If IsNumeric(textTMBM07_1) = False Then
         Cancel = True
         strTit = "資料檢核"
         strMsg = "公報卷期(起)只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_1_GotFocus
      End If
   End If
End Sub

' 公報卷期(迄)
Private Sub textTMBM07_2_LostFocus()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   If IsEmptyText(textTMBM07_2) = False Then
      If IsNumeric(textTMBM07_2) = False Then
         strTit = "資料檢核"
         strMsg = "公報卷期(迄)只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_2_GotFocus
      Else
         If Not ChkRange(textTMBM07_1, textTMBM07_2, "公報卷期") Then
         
         End If
      End If
   End If
End Sub

Private Sub textTMBM07_1_GotFocus()
   InverseTextBox textTMBM07_1
End Sub

Private Sub textTMBM07_2_GotFocus()
   InverseTextBox textTMBM07_2
End Sub

Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False
   If IsEmptyText(textTMBM07_1) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入公報卷期(起)"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM07_1.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(textTMBM07_2) = True Then
      strTit = "資料檢核"
      strMsg = "請輸入公報卷期(迄)"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textTMBM07_2.SetFocus
      GoTo EXITSUB
   End If
   If IsEmptyText(textTMBM07_1) = False And IsEmptyText(textTMBM07_2) = False Then
      If Val(textTMBM07_1) > Val(textTMBM07_2) Then
         strTit = "資料檢核"
         strMsg = "公報卷期範圍不正確"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textTMBM07_1.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Added by Lydia 2017/02/06 檢查公報檔的地區
   If Pub_ChkTMBMValidate(textTMBM07_1, textTMBM07_2) = False Then
      GoTo EXITSUB
   End If
   'end 2017/02/06
   
   CheckDataValid = True
EXITSUB:
End Function


