VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090124 
   Caption         =   "查名人狀態"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   3615
   ScaleWidth      =   4560
   Begin VB.CommandButton Cmd1 
      Caption         =   "回前畫面(&U)"
      Height          =   375
      Index           =   1
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Cmd1 
      Caption         =   "存檔(&O)"
      Height          =   375
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm090124.frx":0000
      Height          =   2385
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   4207
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   8
   End
   Begin VB.Label Label2 
      Caption         =   "請在資料列上點選即可切換查名人狀態"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   360
      TabIndex        =   4
      Top             =   600
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "查名人狀態：N=不分派查名單"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   3360
      Width           =   2535
   End
End
Attribute VB_Name = "frm090124"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/01 改成Form2.0 ; GRD1改字型=新細明體-ExtB
'Memo by Lydia 2015/05/28 GRD1顯示所有記錄,按確定才會回寫記錄到Table，若有修改狀態會寫log記錄。
'Created by Lydia 2015/05/29 新增-查名人狀態
Option Explicit
Dim intLastRow As Integer, intCols As Integer

Public mPreForm As Form
Private Sub Form_Load()
  
   Me.Width = 4560
   Me.Height = 4040
   
   MoveFormToCenter Me
 
   ReadData
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm090124 = Nothing
End Sub

Private Sub Cmd1_Click(Index As Integer)

    Select Case Index
        Case 0 '確定
             If FormSave() = False Then
                MsgBox "請確定是否同時有人在維護查名人員及狀態!", vbCritical
             End If
        Case 1 '回前畫面
    End Select
    mPreForm.Show
    Unload Me
End Sub

Private Function ReadData() As Boolean
   Dim iR As Integer, sqlA As String
   sqlA = "select '' ,TMQSR01,ST02,TMQSR17,TMQM01,TMQM02,TMQM03,nvl(tmqm02||tmqm01,tmqsr01) or2 " & _
          "from TMQSumR a,staff s1,tmqmember where tmqsr01=tmqm01(+) and tmqsr01=s1.st01(+) order by or2 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, sqlA)
   If intI = 1 Then
      ReadData = True
   End If
   Set grd1.Recordset = RsTemp.Clone

   With grd1
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignCenterCenter
      
      .col = 1: .ColWidth(1) = 1500: .Text = "查名/統計人代號"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "姓名"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1000: .Text = "查名人狀態"
      .CellAlignment = flexAlignCenterCenter
      .ColAlignment(.col) = flexAlignCenterCenter
      For iR = 4 To grd1.Cols - 1
         .col = iR: .ColWidth(iR) = 0
      Next
    End With
    
    'Added by Lydia 2018/05/25 檢查狀態
   sqlA = Pub_GetSpecMan("內商查名單分單狀態")
   If UCase(sqlA) = "N" Then
       MsgBox "目前系統設定內商查名單為全面不分單狀態，若有疑問請詢問電腦中心!", vbCritical
   End If
End Function

Private Sub Grd1_Click()
    
    GridClick grd1, intLastRow, 0
    If grd1.TextMatrix(intLastRow, 3) = "" Then
       grd1.TextMatrix(intLastRow, 3) = "N"
    Else
       grd1.TextMatrix(intLastRow, 3) = ""
    End If

End Sub

Private Function FormSave() As Boolean
  Dim rsA As ADODB.Recordset
  Dim strA1 As String, strSNo As String, strStu As String
  Dim idx As Integer, midS As String
On Error GoTo ErrHnd
   
   cnnConnection.BeginTrans
     
    For idx = 1 To grd1.Rows - 1
       strSNo = grd1.TextMatrix(idx, 1)
       strStu = "" & grd1.TextMatrix(idx, 3)
        strA1 = "select tmqsr01,tmqsr17 from tmqsumr where tmqsr01=" & CNULL(strSNo)
        intI = 1
        Set rsA = ClsLawReadRstMsg(intI, strA1)
        If intI = 1 Then
           If Trim(strStu) <> Trim("" & rsA.Fields(1)) Then
              If strStu = "N" Then
                 strA1 = "tmqsr17='N'"
              Else
                 strA1 = "tmqsr17=null"
              End If

              If grd1.TextMatrix(idx, 4) <> grd1.TextMatrix(idx, 5) Then midS = midS & "," & grd1.TextMatrix(idx, 5)
               
              strSql = "update tmqsumr set " & strA1 & " where tmqsr01=" & CNULL(strSNo)
              Pub_SeekTbLog strSql
              cnnConnection.Execute strSql, intI
           End If
        End If
    Next idx
    
    If Len(midS) > 1 Then
        midS = IIf(Left(midS, 1) = ",", Mid(midS, 2, Len(midS) - 1), midS)
        midS = Replace(midS, ",", "','")
      '判斷統計人員是否一起放假
        strA1 = "select tmqm02,nvl(tmqm03,'N') tmqm03,count(*) r1,count(tmqsr17) r2 from tmqmember,tmqsumr " & _
                "where tmqm01<>tmqm02 and tmqm01=tmqsr01(+) and tmqm02 in (" & "'" & midS & "'" & ") group by tmqm02,nvl(tmqm03,'N')"
        If rsA.State <> adStateClosed Then rsA.Close
          Set rsA = New ADODB.Recordset
          rsA.CursorLocation = adUseClient
          rsA.Open strA1, cnnConnection, adOpenStatic, adLockReadOnly
          With rsA
            If rsA.RecordCount > 0 Then
              rsA.MoveFirst
              Do While Not rsA.EOF
                 '情況1:是的話有一位請假，這個代號今天請假,情況2:否的話有一位上班，這個代號今天可排單
                 If (rsA.Fields("tmqm03") = "Y" And rsA.Fields("r2") >= 1) Or _
                    (rsA.Fields("tmqm03") <> "Y" And rsA.Fields("r1") = rsA.Fields("r2")) Then
                    strSql = "update TMQSumR set TMQSR17='N' where TMQSR01=" & CNULL(rsA.Fields(0))
                    Pub_SeekTbLog strSql
                    cnnConnection.Execute strSql, intI
                 End If
                 rsA.MoveNext
              Loop
            End If
          End With
    End If
          
   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrHnd:
   cnnConnection.RollbackTrans
   MsgBox Err.Description
End Function

