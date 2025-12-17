VERSION 5.00
Begin VB.Form Frmacc41k0 
   AutoRedraw      =   -1  'True
   Caption         =   "智權期末結餘保留資料刪除"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   6435
   Begin VB.CommandButton Cmd1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "刪除 xxx年xx月 結餘保留資料"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4530
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   90
      Width           =   1800
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(2)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3330
      TabIndex        =   5
      Top             =   120
      Width           =   1100
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(0)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(1)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "目前結餘保留資料年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   495
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "目前智權點數輸入年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "Frmacc41k0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/12/08 Form2.0已修改 (無需修改)
'Memo by Amy 2022/06/09 MDIChild 設False
'Create by Amy 2021/09/17
Option Explicit

Dim ado41K0 As New ADODB.Recordset
Dim bol0b1HasDt As Boolean, bolHasAx210 As Boolean 'Acc0b1是否有資料/是否已過帳(非轉撥)
Dim strYM As String, strMaxSP01 As String '目前系統年月-1個月/目前智權點數輸入年月
Dim strA0b01 As String, strA0b05 As String '目前過帳日/目前業績輸入關閉年月

Private Sub Cmd1_Click()
    Dim strCmd As String, intRun As Integer
    
    If MsgBox("智權期末結餘保留資料", vbYesNo + vbDefaultButton2) = vbNo Then
        Exit Sub
    End If
    
    strCmd = "Delete From SalesBalance Where SB01=" & Val(strYM) - 191100
    cnnConnection.BeginTrans
    cnnConnection.Execute strCmd, intRun
    If intRun > 0 Then
        Call WirteAxb16(Val(strYM) - 191100, "")
    End If
    cnnConnection.CommitTrans
    Cmd1.Enabled = False
End Sub

Private Sub Form_Load()
    strFormName = Name
    Me.Width = 6675
    Me.Height = 1500
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Call ClearLabel
    
    strYM = Mid(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 1, 6) '系統前一個月
    strA0b01 = GetA0b01(strA0b05)
    strMaxSP01 = GetMaxSP01(True)
    
    Cmd1.Caption = Replace(Cmd1.Caption, "xxx年xx月", Val(Mid(strMaxSP01, 1, 4)) - 1911 & "年" & Right(strMaxSP01, 2) & "月")
    Cmd1.Enabled = False
    Call ShowBt
    
    If Val(strMaxSP01) - 191100 = Val(strA0b05) Then Lbl1(2).Caption = "(已關閉)"
    Lbl1(0).Caption = Val(strMaxSP01) - 191100
    Lbl1(1).Caption = GetMaxSB01
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Call PUB_GetLock("", "Frmacc41k0")
    Set Frmacc41k0 = Nothing
End Sub

Private Sub ShowBt()
    Dim stAxb1(0) As String '結餘是否有修改
    
    Call bolAcc0b1(8, strYM - 191100, stAxb1())
        
    '智權點數已開放但未關閉且結餘資料有修改且SalesBalance資料已產生,且目前結餘作業未在使用中
    If Val(strMaxSP01) - 191100 > Val(strA0b05) And stAxb1(0) = "Y" And TranNoLock("Frmacc41k0", "Frmacc41g0", False) = False _
      And ExistCheck("SalesBalance", "SB01", strMaxSP01 - 191100, strExc(0), False) = True Then
        Cmd1.Enabled = True
    End If
End Sub

Private Function GetMaxSB01() As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer

    strQ = "Select Max(SB01) From SalesBalance "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        GetMaxSB01 = "" & RsQ.Fields(0)
    End If
    RsQ.Close
End Function

Private Sub ClearLabel()
    Dim objLbl As LABEL
    
    For Each objLbl In Lbl1
        objLbl.Caption = ""
    Next
End Sub



