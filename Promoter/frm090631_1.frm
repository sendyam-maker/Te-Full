VERSION 5.00
Begin VB.Form frm090631_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "目標基數複製"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3465
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3465
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1230
      MaxLength       =   5
      TabIndex        =   0
      Top             =   480
      Width           =   930
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1230
      MaxLength       =   5
      TabIndex        =   1
      Top             =   900
      Width           =   900
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   2190
      TabIndex        =   3
      Top             =   30
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1410
      TabIndex        =   2
      Top             =   30
      Width           =   756
   End
   Begin VB.Label Label1 
      Caption         =   "被複製年月："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   5
      Top             =   525
      Width           =   1110
   End
   Begin VB.Label Label1 
      Caption         =   "複製年月："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   4
      Top             =   930
      Width           =   975
   End
End
Attribute VB_Name = "frm090631_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/17 改成Form2.0 (無)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim MyRs As New ADODB.Recordset
Dim MySql As String

Private Sub cmdOK_Click(Index As Integer)
Dim Cancel As Boolean
Select Case Index
Case 0
         Cancel = False
         txt1_Validate 0, Cancel
         If Cancel = True Then Exit Sub
         txt1_Validate 1, Cancel
         If Cancel = True Then Exit Sub
         Set MyRs = New ADODB.Recordset
         '檢查來源月份是否存在
         If MyRs.State = 1 Then MyRs.Close
         MySql = "select * from engradix where er01=" & Mid(ChangeTStringToWString(txt1(0) & "01"), 1, 6) & " "
         MyRs.CursorLocation = adUseClient
         MyRs.Open MySql, cnnConnection, adOpenStatic, adLockReadOnly
         If MyRs.RecordCount = 0 Then
             MsgBox "來源月份沒有基數資料，無法複製！", vbInformation, "警告！"
             Exit Sub
         End If
         '檢查目的月份是否存在
         If MyRs.State = 1 Then MyRs.Close
         MySql = "select * from engradix where er01=" & Mid(ChangeTStringToWString(txt1(1) & "01"), 1, 6) & " "
         MyRs.CursorLocation = adUseClient
         MyRs.Open MySql, cnnConnection, adOpenStatic, adLockReadOnly
         If MyRs.RecordCount > 0 Then
             If MsgBox("目的月份已有基數資料，是否覆蓋？？", vbInformation + vbYesNo, "警告！") = vbNo Then
                Exit Sub
             End If
         End If
         frm090631.TagM = txt1(1)
         frm090631.SrcM = txt1(0)
         frm090631.IsCopy = True
         MsgBox "基數資料複製  及  目標資料產生 作業進行時，畫面將無法運作，" & vbCrLf & "作業完成後，將會有訊息通知完成！"
         Unload Me
Case 1
        Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090631_1 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
     If Len(Trim(txt1(Index))) <> 0 Then
         If IsDate(ChangeTStringToTDateString(txt1(Index) & "01")) = False Then
            MsgBox "年月輸入錯誤", , "USER 輸入錯誤"
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Cancel = True
            Exit Sub
         End If
    End If
End Sub
