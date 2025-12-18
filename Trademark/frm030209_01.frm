VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030209_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "電話通知"
   ClientHeight    =   4500
   ClientLeft      =   135
   ClientTop       =   2415
   ClientWidth     =   8730
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   8730
   Begin VB.CommandButton cmdOK 
      Caption         =   "無期限"
      Height          =   400
      Index           =   1
      Left            =   6645
      TabIndex        =   16
      Top             =   70
      Width           =   975
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3048
      TabIndex        =   4
      Top             =   528
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   960
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "FCT"
      Top             =   576
      Width           =   550
   End
   Begin VB.TextBox Text2 
      Height          =   300
      Left            =   1512
      MaxLength       =   6
      TabIndex        =   1
      Top             =   576
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   300
      Left            =   2352
      MaxLength       =   1
      TabIndex        =   2
      Top             =   576
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   300
      Left            =   2592
      MaxLength       =   2
      TabIndex        =   3
      Top             =   576
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "輸入期限"
      Height          =   400
      Index           =   0
      Left            =   5580
      TabIndex        =   6
      Top             =   70
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7710
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2370
      Left            =   120
      TabIndex        =   5
      Top             =   1950
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4180
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   "V|收 文 日|收 文 號|案　件　性　質|承 辦 人|智權人員|進　　度　　備　　註"
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   1
      Left            =   1920
      TabIndex        =   18
      Top             =   1620
      Width           =   6555
      Size            =   "11562;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   17
      Top             =   1620
      Width           =   855
      Size            =   "1508;450"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label8 
      Caption         =   "申請人1:"
      Height          =   250
      Left            =   120
      TabIndex        =   15
      Top             =   1620
      Width           =   800
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   330
      Left            =   960
      TabIndex        =   14
      Top             =   1215
      Width           =   7575
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "13361;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號:"
      Height          =   250
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   800
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   930
      Width           =   795
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Label4"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   930
      Width           =   2730
   End
   Begin VB.Label Label5 
      Caption         =   "審定號數:"
      Height          =   255
      Left            =   5400
      TabIndex        =   10
      Top             =   930
      Width           =   795
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Label6"
      Height          =   255
      Left            =   6420
      TabIndex        =   9
      Top             =   930
      Width           =   1350
   End
   Begin VB.Label Label7 
      Caption         =   "商標名稱:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1260
      Width           =   795
   End
End
Attribute VB_Name = "frm030209_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2022/07/28
Option Explicit
Dim m_TM(1 To 4) As String
Dim intLastRow As Integer

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click(Index As Integer)
 Dim intJ As Integer, bolChk As Boolean
  
  Me.Tag = ""
  If Trim(lblFM2(0).Caption) = "" Or (m_TM(1) & m_TM(2) & m_TM(3) & m_TM(4) <> Text1 & Text2 & Text3 & Text4) Then
      MsgBox "請先查詢資料!!", vbInformation
      Exit Sub
  End If
  For intJ = 1 To MSHFlexGrid1.Rows - 1
     If MSHFlexGrid1.TextMatrix(intJ, 0) = "v" Then
        bolChk = True
        Me.Tag = "" & MSHFlexGrid1.TextMatrix(intJ, 2)
     End If
  Next
  
  If bolChk = False Then
       MsgBox "請選擇資料 !", vbInformation
       Exit Sub
  End If
         
  If Index = 0 Then '輸入期限
     frm030209_02.Show
     Me.Hide
  ElseIf Index = 1 Then '無期限
     strExc(0) = "請選擇：" & vbCrLf & _
                      "是，已通知代理人，請接著輸入發文日" & vbCrLf & _
                      "否，無需通知代理人更新發文日=111111" & vbCrLf & _
                      "取消，中斷作業。"
     intJ = MsgBox(strExc(0), vbInformation + vbYesNoCancel + vbDefaultButton3)

     If intJ = vbYes Then 'Yes=已通知代理人
JumpReInput:
         strExc(1) = UCase(InputBox("已通知代理人，請接著輸入發文日或是取消：" & vbCrLf & "P.S.發文日不可大於系統日", "已通知代理人", strSrvDate(2)))
         If strExc(1) = "" Then
             Exit Sub
         Else
             If Len(strExc(1)) <> 7 Then
                 GoTo JumpReInput
             Else
                If strExc(1) > strSrvDate(2) Then
                    GoTo JumpReInput
                Else
                    If ChkDate(strExc(1)) = False Then
                        GoTo JumpReInput
                    End If
                End If
             End If
             strSql = "Update CaseProgress set cp27=" & DBDATE(strExc(1)) & " where cp09='" & Me.Tag & "' "
             cnnConnection.Execute strSql
         End If
     ElseIf intJ = vbNo Then 'No=無需通知代理人
         strSql = "Update CaseProgress set cp27='19221111' where cp09='" & Me.Tag & "' "
         cnnConnection.Execute strSql
     End If
     If intJ <> vbCancel Then
         ClearForm
     End If
  End If
  
End Sub

Private Sub cmdQuery_Click()
   
   If Trim(Text1) = "" Then
       MsgBox "請輸入本所案號!!", vbCritical
       Text1.SetFocus
       Text1_GotFocus
       Exit Sub
   End If
   If Trim(Text2) = "" Then
       MsgBox "請輸入本所案號!!", vbCritical
       Text2.SetFocus
       Text2_GotFocus
       Exit Sub
   End If
   
   Label4 = ""
   Label6 = ""
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   m_TM(1) = Text1
   m_TM(2) = Text2
   m_TM(3) = Text3
   m_TM(4) = Text4
   Combo1.Clear
   InitGrid 7, MSHFlexGrid1
   GridHead
   
   strExc(0) = "select tm05,tm06,tm07,tm23, nvl(cu05,nvl(cu04,cu06)) cname1,tm12,tm15 " & _
                     "From trademark, customer where substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) " & _
                     "and tm01='" & m_TM(1) & "' and tm02='" & m_TM(2) & "' and tm03='" & m_TM(3) & "' and tm04='" & m_TM(4) & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
        MsgBox "查無本所案號!!", vbCritical
        Exit Sub
   Else
        lblFM2(0) = "" & RsTemp.Fields("tm23")
        lblFM2(1) = "" & RsTemp.Fields("cname1")
        Label4 = "" & RsTemp.Fields("tm12")
        Label6 = "" & RsTemp.Fields("tm15")
        If "" & RsTemp.Fields("tm07") <> "" Then
            Combo1.AddItem "外：" & RsTemp.Fields("tm07"), 0
        End If
        If "" & RsTemp.Fields("tm06") <> "" Then
            Combo1.AddItem "英：" & RsTemp.Fields("tm06"), 0
        End If
        Combo1.AddItem "中：" & RsTemp.Fields("tm05"), 0
        Combo1.ListIndex = 0
   End If
   strExc(0) = "select '', sqldatet(cp05) CP05T ,cp09,cpm03,staff.st02 as st1,staff1.st02 as st2,cp64" & _
          " from caseprogress, casepropertymap,staff,staff staff1" & _
          " where " & ChgCaseprogress(m_TM(1) & m_TM(2) & m_TM(3) & m_TM(4)) & _
          " and cp09<'D' and cp10='1727' and cp01=cpm01(+) and cp10=cpm02(+)" & _
          " and cp14=staff.st01(+) and cp13=staff1.st01(+)" & _
          " and cp27 is null and cp57 is null" & _
          " order by CP05 desc"
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI <> 2 Then Set MSHFlexGrid1.Recordset = RsTemp
   GridHead
   
   '若只搜尋到一筆時直接勾選
   If Me.MSHFlexGrid1.Rows = 2 Then
      MSHFlexGrid1_Click
   End If
      
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Label4 = ""
   Label6 = ""
   InitGrid 7, MSHFlexGrid1
   GridHead
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030209_01 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1 <> "FCT" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub GridHead()
Dim intJ As Integer
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      .Visible = False
      .row = 0
      .col = 0: .ColWidth(0) = 200: .Text = "v"
      .CellAlignment = flexAlignCenterCenter
      .col = 1: .ColWidth(1) = 1000: .Text = "收文日"
      .CellAlignment = flexAlignCenterCenter
      .col = 2: .ColWidth(2) = 1000: .Text = "收文號"
      .CellAlignment = flexAlignCenterCenter
      .col = 3: .ColWidth(3) = 1200: .Text = "案件性質"
      .CellAlignment = flexAlignCenterCenter
      .col = 4: .ColWidth(4) = 1000: .Text = "承辦人"
      .CellAlignment = flexAlignCenterCenter
      .col = 5: .ColWidth(5) = 1000: .Text = "智權人員"
      .CellAlignment = flexAlignCenterCenter
      .col = 6: .ColWidth(6) = 2400: .Text = "進度備註"
      If .Cols > 7 Then
          For intJ = 7 To .Cols
              .col = intJ
              .ColWidth(intJ) = 0
          Next intJ
      End If
      .Visible = True
      If .Rows > 1 Then .row = 1
   End With
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

Private Sub Text3_GotFocus()
  TextInverse Text3
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub

Public Sub ClearForm()
   '保留原輸入的系統類別
   Text2 = Empty
   Text3 = Empty
   Text4 = Empty
   Label4 = Empty
   Label6 = Empty
   lblFM2(0).Caption = Empty
   lblFM2(1).Caption = Empty
   Combo1.Clear
   InitGrid 7, MSHFlexGrid1
   GridHead
   
End Sub

