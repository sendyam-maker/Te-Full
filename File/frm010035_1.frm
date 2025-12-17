VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010035_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "借閱資料維護"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   1935
   ClientWidth     =   6270
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6270
   Begin VB.CommandButton cmdExit 
      Caption         =   "回上一頁"
      Height          =   400
      Left            =   5220
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   10
      Width           =   1035
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定"
      Default         =   -1  'True
      Height          =   400
      Left            =   4380
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   10
      Width           =   800
   End
   Begin MSForms.TextBox Text1 
      Height          =   1080
      Left            =   1080
      TabIndex        =   23
      Top             =   2880
      Width           =   5000
      VariousPropertyBits=   -1466939365
      ScrollBars      =   3
      Size            =   "8819;1905"
      Value           =   "text1"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   8
      Left            =   1005
      TabIndex        =   22
      Top             =   2160
      Width           =   1200
      VariousPropertyBits=   27
      Caption         =   "LblBK(8)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "譯        者："
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   1005
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   20
      Top             =   1800
      Width           =   5000
      VariousPropertyBits=   27
      Caption         =   "LblBK(7)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   6
      Left            =   1080
      TabIndex        =   19
      Top             =   1500
      Width           =   5000
      VariousPropertyBits=   27
      Caption         =   "LblBK(6)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   18
      Top             =   1200
      Width           =   5000
      VariousPropertyBits=   27
      Caption         =   "LblBK(5)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   17
      Top             =   900
      Width           =   5000
      VariousPropertyBits=   27
      Caption         =   "LblBK(4)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   16
      Top             =   540
      Width           =   1020
      VariousPropertyBits=   27
      Caption         =   "LblBK(3)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   2
      Left            =   3060
      TabIndex        =   15
      Top             =   240
      Width           =   1300
      VariousPropertyBits=   27
      Caption         =   "LblBK(2)"
      Size            =   "2293;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   14
      Top             =   240
      Width           =   1020
      VariousPropertyBits=   27
      Caption         =   "LblBK(1)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   10
      Left            =   1680
      TabIndex        =   13
      Top             =   2460
      Width           =   1200
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Caption         =   "LblBK(10)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   11
      Left            =   4245
      TabIndex        =   10
      Top             =   2460
      Width           =   1200
      BackColor       =   -2147483638
      VariousPropertyBits=   27
      Caption         =   "LblBK(11)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "保管單位："
      Height          =   255
      Index           =   12
      Left            =   3360
      TabIndex        =   9
      Top             =   2460
      Width           =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "保管人員(聯絡人)："
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   8
      Top             =   2460
      Width           =   1800
   End
   Begin VB.Label Label1 
      Caption         =   "作者 (英)："
      Height          =   255
      Index           =   8
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "作者 (中)："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "書名 (英)："
      Height          =   255
      Index           =   9
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "備　　註："
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ＩＳＢＮ："
      Height          =   255
      Index           =   1
      Left            =   2100
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "書名 (中)："
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   2
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "類　　別："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "圖書編號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
End
Attribute VB_Name = "frm010035_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/27 Form2.0已修改 LblBK/Text1
'2016/10/03 Create by Amy
Option Explicit

Public strPreRow As String '前畫面列 for 更新前畫面狀態用
Dim i As Integer

Private Sub cmdExit_Click()
    tmpBol = fnCancelNowFormAndShowParentForm(Me)
End Sub

Private Sub cmdok_Click()
    Dim strTo As String, strContent As String
    
    If InsLoanRecord = True Then
        strTo = Mid(LblBK(10), 1, InStr(LblBK(10), " ") - 1)
        If LblBK(4) <> MsgText(601) Then strContent = LblBK(4)
        If LblBK(5) <> MsgText(601) Then
            If strContent <> MsgText(601) Then
                strContent = strContent & "(" & LblBK(5) & ")"
            Else
                strContent = LblBK(5)
            End If
        End If
        strContent = "圖書編號：" & LblBK(1) & vbCrLf & _
                            "ＩＳＢＮ：" & LblBK(2) & vbCrLf
        If LblBK(4) <> MsgText(601) And LblBK(5) <> MsgText(601) Then
            strContent = strContent & "書　　名：" & LblBK(4) & "(" & LblBK(5) & ")" & vbCrLf
        ElseIf LblBK(4) <> MsgText(601) Then
            strContent = strContent & "書　　名：" & LblBK(4) & vbCrLf
        Else
            strContent = strContent & "書　　名：" & LblBK(5) & vbCrLf
        End If
        If LblBK(6) <> MsgText(601) And LblBK(7) <> MsgText(601) Then
            strContent = strContent & "作　　者：" & LblBK(6) & "(" & LblBK(7) & ")" & vbCrLf
        ElseIf LblBK(6) <> MsgText(601) Then
            strContent = strContent & "作　　者：" & LblBK(6) & vbCrLf
        Else
            strContent = strContent & "作　　者：" & LblBK(7) & vbCrLf
        End If
               
        If Trim(Text1) <> MsgText(601) Then strContent = strContent & vbCrLf & vbCrLf & "借閱備註：" & vbCrLf & Text1
        If strTo = MsgText(601) Then
            PUB_SendMail strUserNum, "A2004", "", "系統發出之圖書借閱申請訊息有誤", strContent
        Else
            PUB_SendMail strUserNum, strTo, "", "圖書借閱申請，請至一般作業－＞圖書借閱資料查詢　之個人記錄 確認！", strContent
        End If
        frm010035.grdDataList.TextMatrix(strPreRow, 8) = "借閱申請中"
        frm010035.grdDataList.TextMatrix(strPreRow, 9) = GetPrjSalesNM(strUserNum)
        cmdExit_Click
    End If
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

Public Function QueryRecord(ByVal stBK01 As String) As Boolean
    Dim Lbl 'Modify by Amy 2021/07/27  原:Lbl As LABEL
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strTmp As String
    Dim idx As Integer
    
    QueryRecord = False
  
    strQ = "Select * From BooksData Where BK01='" & stBK01 & "' "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        ClearField
        With RsQ
            For Each Lbl In LblBK
                idx = Lbl.Index
                'Modify by Amy 2021/07/27 +.Caption
                Lbl.Caption = "" & RsQ.Fields("BK" & Format(idx, "00"))
                If idx = 3 Then Lbl.Caption = "" & GetBK03(Lbl)
                If idx = 10 Then Lbl.Caption = Lbl.Caption & " " & StaffQuery(Lbl.Caption)
            Next
        End With
        QueryRecord = True
       
    End If
    RsQ.Close
End Function

Private Sub ClearField()
    Dim Lbl 'Modify by Amy 2021/07/28 原:Lbl As LABEL
    
    For Each Lbl In LblBK
        Lbl.Caption = Empty
    Next
    Text1 = Empty
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm010035_1 = Nothing
End Sub

Private Function InsLoanRecord() As Boolean
    Dim strExe As String

On Error GoTo ErrHnd
    
    InsLoanRecord = False
    strExe = GetSerialNo_Lib(1)
    strExe = "Insert Into LoanRecord (LR01,LR02,LR03,LR04,LR07,LR08,LR09,LR10) Values(" & _
                    CNULL(strExe) & ",'1'," & CNULL(LblBK(1)) & "," & strSrvDate(1) & "," & _
                    CNULL(Text1) & "," & CNULL(strUserNum) & "," & strSrvDate(1) & "," & ChgSQL(Left(Format(ServerTime, "000000"), 4)) & ")"
    cnnConnection.Execute strExe
    InsLoanRecord = True
    Exit Function
    
ErrHnd:
   MsgBox "程式有誤請洽電腦中心！" & vbCrLf & Err.Description
End Function

Private Function GetBK03(ByVal stVal As String) As String
    GetBK03 = ""
    Select Case stVal
        Case 1
            GetBK03 = stVal & ".專利"
        Case 2
            GetBK03 = stVal & ".商標"
        Case 3
            GetBK03 = stVal & ".法律"
        Case 4
            GetBK03 = stVal & ".電腦"
        Case 5
            GetBK03 = stVal & ".其他"
    End Select
End Function

Private Sub Text1_GotFocus()
    OpenIme
End Sub
