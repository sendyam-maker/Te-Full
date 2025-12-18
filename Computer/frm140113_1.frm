VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140113_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "教育訓練登錄作業-議題"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5085
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5085
   StartUpPosition =   3  '系統預設值
   Begin VB.OptionButton Opt1 
      Caption         =   "外賓"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   4170
      TabIndex        =   17
      Top             =   1590
      Width           =   800
   End
   Begin VB.OptionButton Opt1 
      Caption         =   "員工"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   3060
      TabIndex        =   16
      Top             =   1590
      Value           =   -1  'True
      Width           =   900
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "刪除 ->"
      Height          =   285
      Left            =   2910
      TabIndex        =   14
      Top             =   2220
      Width           =   700
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "<- 新增"
      Height          =   285
      Left            =   2910
      TabIndex        =   13
      Top             =   1860
      Width           =   700
   End
   Begin VB.CommandButton Command1 
      Caption         =   "參加人員"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   360
      TabIndex        =   11
      Top             =   2610
      Width           =   1140
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   4050
      TabIndex        =   10
      Top             =   2610
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2970
      TabIndex        =   9
      Top             =   2610
      Width           =   1005
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   1
      Left            =   675
      TabIndex        =   3
      Top             =   930
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm"
      Format          =   150929411
      UpDown          =   -1  'True
      CurrentDate     =   40942
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   315
      Index           =   2
      Left            =   1995
      TabIndex        =   5
      Top             =   930
      Width           =   960
      _ExtentX        =   1693
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm"
      Format          =   150929411
      UpDown          =   -1  'True
      CurrentDate     =   40942
   End
   Begin MSForms.ListBox lstSpeaker 
      Height          =   960
      Left            =   30
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1590
      Width           =   2800
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "4939;1693"
      MatchEntry      =   0
      MultiSelect     =   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text2 
      Height          =   315
      Left            =   3645
      TabIndex        =   7
      Top             =   1845
      Width           =   1320
      VariousPropertyBits=   679493659
      Size            =   "2328;556"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   855
      Left            =   675
      TabIndex        =   1
      Top             =   30
      Width           =   4350
      VariousPropertyBits=   -1467989989
      ScrollBars      =   2
      Size            =   "7673;1508"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   0
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblNo 
      AutoSize        =   -1  'True
      Caption         =   "LblNo"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   3690
      TabIndex        =   15
      Top             =   2220
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "( 可輸入員工編號或名稱)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3105
      TabIndex        =   8
      Top             =   1350
      Width           =   1785
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主講人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   6
      Top             =   1350
      Width           =   780
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   " ～"
      Height          =   180
      Left            =   1680
      TabIndex        =   4
      Top             =   990
      Width           =   270
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "時間："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   60
      TabIndex        =   2
      Top             =   990
      Width           =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "議題："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   585
   End
End
Attribute VB_Name = "frm140113_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/01/06 Form2.0已修改 text1/text2/lstSpeaker
'Memo By Sonia 2012/12/6 智權人員欄已修改
'Created by Morgan 2012/2/13
Option Explicit

Public m_curRow As Integer
Public m_strMode As String
Public m_stNumList As String
Public bolPublic As Boolean 'Add by Amy 2018/09/18 是公開
Public m_DeptNo As String 'Add by Amy 2019/11/12 目前登入者部門
Dim Old_Speaker, bolChkOnly As Boolean, strMsg As String 'Add by Amy 2020/12/28 原始演講者/只Check/錯誤訊息
 
'Add by Amy 2020/12/28
Private Sub cmdAdd_Click()
    Dim i As Integer, bolCancel As Boolean
    Dim strData As String, strChk As String
    
    If Trim(Text2) = MsgText(601) Then Exit Sub
    
    bolChkOnly = True
    Call Text2_LostFocus
    If strMsg <> MsgText(601) Then
        strMsg = ""
        Exit Sub
    End If
    bolChkOnly = False
    
    strMsg = "": strChk = LblNo
    If strChk <> MsgText(601) Then
        strChk = Text2 & "(" & strChk & ")"
    Else
        strChk = Text2
    End If
    For i = 0 To lstSpeaker.ListCount - 1
        strData = lstSpeaker.List(i)
        If strData = strChk Then
            strMsg = lstSpeaker.List(i) & "資料重覆！"
        End If
    Next i
    If strMsg <> MsgText(601) Then
        MsgBox strMsg
        strMsg = ""
        Exit Sub
    End If
    
    strData = Text2
    If LblNo <> MsgText(601) Then strData = strData & "(" & LblNo & ")"
    lstSpeaker.AddItem strData, 0
    Text2 = ""
    LblNo = ""
End Sub

Private Sub cmdDel_Click()
    If lstSpeaker.ListCount > 0 Then
        lstSpeaker.RemoveItem lstSpeaker.ListIndex
    End If
End Sub
'end 2020/12/28

Private Sub Command1_Click(Index As Integer)
   Dim stLblMsg As String 'Add by Amy 2020/11/27
   
   Select Case Index
      '確定
      Case 0
         'Add by Amy 2022/01/06
         '檢查畫面的 TextBox是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "TextBox") = False Then
            Exit Sub
         End If
         'bug-無議題時按下 frm140113的修改鈕會錯
         If Trim(Text1.Text) = MsgText(601) Then
            MsgBox "議題不可為空"
            Exit Sub
         End If
         UpdateGrid
         Unload Me
      '取消
      Case 1
         Unload Me
      '參加人員
      Case 2
         'Add by Amy 2020/11/27 避免User忘記,彈提醒
         stLblMsg = "需參加之人員都需加入「已選人員」，才會發信通知！"
         stLblMsg = stLblMsg & vbCrLf & "「登記」頁籤及印「點名簿」才會出現。"
         'end 2020/11/27
         
         Set frm140113_2.fmParent = Me
         frm140113_2.strDeptNo = m_DeptNo 'Add by Amy 2019/11/12
         frm140113_2.bolPublic = bolPublic 'Modify by Amy 2018/09/18
         frm140113_2.m_stNumList = m_stNumList
         frm140113_2.Caption = frm140113_2.Caption & "(參加人員)"
         frm140113_2.lblMemo = stLblMsg 'Add by Amy 2020/11/27
         frm140113_2.Show vbModal
   End Select
   
End Sub

Private Function UpdateGrid() As String
   Dim strSeaker As String 'Add by Amy 2020/12/28
   
   With frm140113.MSHFlexGrid1
      If m_strMode = "A" Then
         If .TextMatrix(.Rows - 1, 0) <> "" Then .Rows = .Rows + 1
         m_curRow = .Rows - 1
      End If
      .TextMatrix(m_curRow, 1) = Text1
      .TextMatrix(m_curRow, 2) = Format(DTPicker1(1).Value, "HHmm")
      .TextMatrix(m_curRow, 3) = Format(DTPicker1(2).Value, "HHmm")
      'Modify by Amy 2020/12/28 主講者多筆
      '.TextMatrix(m_curRow, 4) = Text2
      Call SetFrm140113Grid1Speaker
      'end 2020/12/28
      frm140113.GridRefresh False, m_curRow
   End With
   frm140113.SetBookInList m_curRow, m_stNumList
   
End Function

Private Sub Form_Load()
   MoveFormToCenter Me, True
   LblNo = "" 'Add by Amy 2020/12/28
   'Modify by Amy 2022/01/06一開始將ListBox拉到需要的大小,字型會自動放大；所以畫面預設為一列高度,Form_Load才放大到需要的大小
    lstSpeaker.Height = 960
    lstSpeaker.Width = 2800
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_DeptNo = "" 'Add by Amy 2019/11/12
    
   Set frm140113_1 = Nothing
End Sub

'Add by Amy 2020/12/28
'因同名同姓只有一筆時無法知道人員為所內員工或外賓,故加選項
Private Sub Opt1_Click(Index As Integer)
    Dim bolCancel As Boolean
    
    If Index = 1 And opt1(1).Value = True Then
        LblNo = ""
    ElseIf Trim(Text2) <> MsgText(601) Then
        Call Text2_LostFocus
    End If
End Sub

'員編及姓名都檢查(不寫於Validate,因一開始選員工,但輸入姓名非員工會一直彈訊息,無法切至外賓)
Private Sub Text2_LostFocus()
    Dim stQ As String
   
    strMsg = ""
    If (m_strMode = "A" Or m_strMode = "E") And Trim(Text2) <> MsgText(601) Then
        If bolChkOnly = False Then LblNo = "" '輸A4023跳離已將員編寫至LblNo,再按「新增鈕」再檢查會LblNo被清空
        stQ = "Select st01,st02,st04 From Staff Where st04='1' And st01>'6' and st01<'F' and SubStr(st01,1,3)<>'999' And SubStr(st01,4,1)<>'9' "
        '員編查
        If Left(Text2, 1) < "z" And (Len(Text2) = 5 Or Len(Text2) = 6) Then
            strExc(0) = stQ & " And st01='" & UCase(Text2) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                If "" & RsTemp.Fields("st04") = "2" Then
                    strMsg = "員工已離職，請確認！"
                End If
                If bolChkOnly = False Then
                    LblNo = UCase(Text2)
                    Text2 = "" & RsTemp.Fields("st02")
                End If
            ElseIf opt1(0).Value = True Then
                strMsg = "無此員工，請確認！"
            End If
        End If
        If strMsg <> MsgText(601) Then
            MsgBox strMsg
            Exit Sub
        End If
        If LblNo <> MsgText(601) Then Exit Sub
        
        If opt1(0).Value = True Then
            '姓名查
            strExc(0) = stQ & " And st02='" & UCase(Text2) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
                If RsTemp.RecordCount > 1 Then
                    strMsg = "同名同姓有 " & RsTemp.RecordCount & " 筆" & vbCrLf & _
                                "請至於「共同查詢」 查員編後輸入員編 ！"
                Else
                    If "" & RsTemp.Fields("st04") = "2" Then
                        strMsg = "員工已離職，請確認！"
                    End If
                    If bolChkOnly = False Then LblNo = "" & UCase(RsTemp.Fields("st01"))
                End If
            Else
                strMsg = "無此員工，請確認！"
            End If
            If strMsg <> MsgText(601) Then
                MsgBox strMsg
                Exit Sub
            End If
        End If
        
    End If
End Sub

'Mark by Amy 2020/12/28 改寫至Text2_LostFocus
'Private Sub Text2_Change()
'   If Left(Text2, 1) < "z" And (Len(Text2) = 5 Or Len(Text2) = 6) Then
'      strExc(0) = "select st02 from staff where st01='" & UCase(Text2) & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         Text2 = "" & RsTemp.Fields(0)
'      End If
'   End If
'End Sub

'Add by Amy 2020/12/28 演講者多筆
Public Sub SetSS06List(ByVal stSpeaker As String)
    Dim ii As Integer
   
    lstSpeaker.Clear
    
    Old_Speaker = Split(stSpeaker, ";")
    For ii = LBound(Old_Speaker) To UBound(Old_Speaker)
        lstSpeaker.AddItem Old_Speaker(ii), 0
    Next ii
    
End Sub

Private Sub SetFrm140113Grid1Speaker()
    Dim ii As Integer, stTmp As String, stTmp2 As String
    Dim stDisplay As String, stOrgData As String
    
    If lstSpeaker.ListCount = 0 Then Exit Sub
    
    For ii = lstSpeaker.ListCount - 1 To 0 Step -1
        stTmp = lstSpeaker.List(ii)
        stTmp2 = stTmp
        If InStr(stTmp, "(") Then
            stTmp = Mid(stTmp, 1, Val(InStr(stTmp, "(")) - 1)
            stTmp2 = Mid(stTmp2, Val(InStr(stTmp2, "(")) + 1)
            stTmp2 = Replace(stTmp2, ")", "")
        End If
        stDisplay = stDisplay & "," & stTmp
        stOrgData = stOrgData & ";" & stTmp2
    Next ii
    frm140113.MSHFlexGrid1.TextMatrix(m_curRow, 4) = Mid(stDisplay, 2)
    frm140113.MSHFlexGrid1.TextMatrix(m_curRow, 5) = Mid(stOrgData, 2)
End Sub
'end 2020/12/28
