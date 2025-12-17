VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010036 
   BorderStyle     =   1  '單線固定
   Caption         =   "圖書借閱資料查詢(檔案室使用)"
   ClientHeight    =   5100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5100
   ScaleWidth      =   6945
   Begin VB.ComboBox cboStatus 
      Height          =   300
      ItemData        =   "frm010036.frx":0000
      Left            =   1080
      List            =   "frm010036.frx":0013
      TabIndex        =   6
      Text            =   "cboStatus"
      Top             =   1630
      Width           =   1680
   End
   Begin VB.CheckBox Check1 
      Caption         =   "含已遺失或銷毀"
      Height          =   300
      Left            =   3000
      TabIndex        =   10
      Top             =   1635
      Width           =   1800
   End
   Begin VB.ComboBox cboClass 
      Height          =   300
      ItemData        =   "frm010036.frx":0033
      Left            =   1080
      List            =   "frm010036.frx":0049
      TabIndex        =   5
      Text            =   "cboClass"
      Top             =   1300
      Width           =   1680
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "書籍資料"
      Height          =   350
      Index           =   0
      Left            =   4920
      Style           =   1  '圖片外觀
      TabIndex        =   29
      Top             =   1620
      Width           =   950
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "借閱記錄"
      Height          =   350
      Index           =   1
      Left            =   5910
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   1620
      Width           =   950
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   2
      Left            =   5280
      MaxLength       =   7
      TabIndex        =   9
      Top             =   2020
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   1
      Left            =   3960
      MaxLength       =   7
      TabIndex        =   8
      Top             =   2020
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Index           =   0
      Left            =   1080
      MaxLength       =   5
      TabIndex        =   7
      Text            =   "A2004"
      Top             =   2020
      Width           =   750
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   350
      Left            =   6060
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      Default         =   -1  'True
      Height          =   350
      Left            =   5160
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   45
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2700
      Left            =   120
      TabIndex        =   13
      Top             =   2350
      Width           =   6745
      _ExtentX        =   11906
      _ExtentY        =   4763
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V| 編號|書名|ISBN|作者|譯者|類別|保管人|狀態|借閱人|借閱日|上架日|出刊日"
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
      _Band(0).Cols   =   13
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   1000
      Width           =   4850
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   90
      Width           =   500
      VariousPropertyBits=   679495707
      MaxLength       =   4
      Size            =   "882;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   420
      Width           =   4845
      VariousPropertyBits=   679495706
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   1
      Left            =   3360
      TabIndex        =   1
      Top             =   90
      Width           =   1700
      VariousPropertyBits=   679495707
      MaxLength       =   20
      Size            =   "2999;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   300
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   4850
      VariousPropertyBits=   679495707
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "狀　　態："
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   30
      Top             =   1630
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(模糊比對)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   11
      Left            =   6000
      TabIndex        =   27
      Top             =   1000
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "譯　　者："
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   26
      Top             =   1000
      Width           =   1095
   End
   Begin VB.Label LblEmp 
      BackColor       =   &H8000000A&
      Height          =   255
      Left            =   1845
      TabIndex        =   25
      Top             =   2020
      Width           =   960
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(模糊比對)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   6000
      TabIndex        =   24
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(模糊比對)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   5
      Left            =   6000
      TabIndex        =   23
      Top             =   430
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   8
      Left            =   4980
      TabIndex        =   22
      Top             =   2080
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "借閱日期："
      Height          =   255
      Index           =   6
      Left            =   3000
      TabIndex        =   21
      Top             =   2020
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "圖書編號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   20
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "類　　別："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   1300
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "書　　名："
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   430
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ＩＳＢＮ："
      Height          =   255
      Index           =   1
      Left            =   2505
      TabIndex        =   17
      Top             =   90
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "(流水號)"
      Height          =   255
      Left            =   1605
      TabIndex        =   16
      Top             =   90
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "作　　者："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "借閱人員："
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   14
      Top             =   2020
      Width           =   1095
   End
End
Attribute VB_Name = "frm010036"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/27 Form2.0已修改 Text1/GrdDataList
'2016/10/03 Create by Amy
Option Explicit

Dim i As Integer
Dim arrField, intWidth
Dim cmdState As Integer

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 '設只能選
End Sub

Private Sub cboStatus_Click()
    Check1.Value = 0
    If cboStatus = "遺失" Or cboStatus = "銷毀" Then
        Check1.Value = 1
    End If
End Sub

Private Sub cboStatus_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 '設只能選
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
    cmdState = Index
    PubShowNextData
End Sub

Private Sub cmdSearch_Click()
    Dim oTxt 'Modify by Amy 2021/07/27 改From2.0 拿掉 As TextBox
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strBK As String, strLR As String, strWhere As String
    Dim intChoose  As Integer
        
    If TxtValidate = False Then MsgBox "請輸入查詢條件！", vbInformation: Exit Sub
    
    intChoose = 0
    For Each oTxt In Text1
        If oTxt <> MsgText(601) Then
            Select Case oTxt.Index
                Case 0
                    strBK = strBK & " And BK01='" & oTxt & "' "
                Case 1
                    strBK = strBK & " And BK02='" & oTxt & "' "
                Case 2
                    strBK = strBK & " And (Upper(BK04) like '%" & UCase(oTxt) & "%' Or Upper(BK05) like '%" & UCase(oTxt) & "%')"
                Case 3
                    strBK = strBK & " And (Upper(BK06) like '%" & UCase(oTxt) & "%' Or Upper(BK07) like '%" & UCase(oTxt) & "%')"
                Case 4
                    strBK = strBK & " And Upper(BK08) like '%" & UCase(oTxt) & "%'"
            End Select
        End If
    Next
    If cboClass <> MsgText(601) Then strBK = strBK & " And BK03=" & Mid(cboClass, 1, InStr(cboClass, ".") - 1)
    If cboStatus <> MsgText(601) Then
        Select Case cboStatus
            Case "可借閱"
                strWhere = strWhere & " And LR02='X'"
            Case "借閱"
                strWhere = strWhere & " And LR02>='1' And LR02<='W'"
            Case "遺失"
                strWhere = strWhere & " And LR02='Y'"
            Case "銷毀"
                strWhere = strWhere & " And LR02='Z'"
        End Select
    End If
    
    '若有下借閱人員或借閱日期條件,顯示條件內的借閱記錄
    If Trim(Text2(0)) & Trim(Text2(1)) & Trim(Text2(2)) <> MsgText(601) Then
        intChoose = 1
        If Text2(0) <> MsgText(601) Then strLR = strLR & " And LR08='" & Text2(0) & "' "
        If Text2(1) <> MsgText(601) Then strLR = strLR & " And LR04>=" & Val(Text2(1)) + 19110000
        If Text2(2) <> MsgText(601) Then strLR = strLR & " And LR04<=" & Val(Text2(2)) + 19110000
    '只查狀態
    ElseIf strBK = MsgText(601) And strWhere <> MsgText(601) Then
        intChoose = 2
        strLR = "Select * From LoanRecord a Where LR01||LR02=(Select Max(LR01||LR02) From LoanRecord  Where a.LR03=LR03 )"
    '依條件只顯示最新一筆借閱記錄
    Else
        strBK = "Select * From BooksData Where " & Mid(strBK, 5)
        strLR = "Select * From LoanRecord a Where LR01||LR02=(Select Max(LR01||LR02) From LoanRecord  Where a.LR03=LR03 )" & strLR
    End If

    strQ = "Select '' as V,BK01 as 編號,Decode(BK04,null,BK05,Decode(BK05,null,BK04,BK04||'('||BK05||')')) as 書名,Nvl(BK02,'') as ISBN," & _
            "Decode(BK06,null,BK07,Decode(BK07,null,BK06,BK06||'('||BK07||')')) as 作者,Nvl(BK08,'') as 譯者," & _
            "Decode(BK03,'1','專利','2','商標','3','法律','4','電腦','5','其他') as 類別,k.ST02 as 保管人," & _
            "Decode(LR06,null,Decode(LR02, '1','借閱申請中', 'X','可借閱', 'Y','遺失', 'Z','銷毀', '延期申請中'), Decode(LR02, 'X','可借閱', 'Y','遺失', 'Z','銷毀', '借閱中')) as  狀態," & _
            "Decode(LR02,'Z','',b.ST02) as 借閱人,Decode(LR02,'Z','',Decode(LR06,null,'',sqldatet(LR04))) as 借閱日," & _
            "sqldatet(BK09) as 上架日,Decode(BK13,null,'',sqldatet(BK13)) as 出刊日, " & _
            "Nvl(LR06,'') as LR06,BK10,LR01,LR02,LR08 "
    '依條件只顯示最新一筆借閱記錄
    If intChoose = 0 Then
        strQ = strQ & "From (" & strBK & "),(" & strLR & ")" & _
                ",Staff k,Staff b Where BK01=LR03(+) And BK10=k.ST01(+) And LR08=b.ST01(+) " & _
                IIf(Check1.Value = 0, " And LR02<>'Y' And LR02<>'Z' ", "") & strWhere & _
                " Order by LR03,LR01,LR02"
    '若有下借閱人員或借閱日期條件,顯示條件內的借閱記錄
    ElseIf intChoose = 1 Then
        strQ = strQ & "From BooksData,LoanRecord" & _
                ",Staff k,Staff b Where BK01=LR03(+) And BK10=k.ST01(+) And LR08=b.ST01(+) " & _
                IIf(Check1.Value = 0, " And LR02<>'Y' And LR02<>'Z' ", "") & strWhere & strBK & strLR & _
                " Order by LR03,LR01,LR02"
    '只查狀態
    Else
        strQ = strQ & "From BooksData,(" & strLR & ")" & _
                ",Staff k,Staff b Where BK01=LR03(+) And BK10=k.ST01(+) And LR08=b.ST01(+) " & _
                strWhere & " Order by LR03,LR01,LR02"
    End If
    
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    grdDataList.Clear
    grdDataList.FixedCols = 0
    If RsQ.RecordCount <> 0 Then
        Set grdDataList.Recordset = RsQ
        SetGridWidth
    Else
        grdDataList.Rows = 2
        SetGridWidth
        MsgBox "查無資料！", vbInformation
    End If
    grdDataList.FixedCols = 3
    grdDataList.ColAlignmentFixed(GetValue("編號")) = flexAlignLeftCenter
    grdDataList.ColAlignmentFixed(GetValue("書名")) = flexAlignLeftCenter
    RsQ.Close
    Set RsQ = Nothing
End Sub

Private Sub Form_Load()
    ReDim arrField(17)
    ReDim intWidth(17)
    arrField = Array("V", "編號", "書名", "ISBN", "作者", "譯者", "類別", "保管人", "狀態", "借閱人", "借閱/延期日", "上架日", "出刊日", _
                                "LR06", "BK10", "LR01", "LR02", "LR08")
    intWidth = Array(200, 500, 1300, 1000, 1000, 1000, 500, 700, 1000, 700, 1000, 1000, 1000, _
                                0, 0, 0, 0, 0)
                                
    MoveFormToCenter Me
    ClearField
    SetGridWidth
    grdDataList.FixedCols = 3 '固定前3欄
    cboClass = ""
End Sub

Private Function TxtValidate() As Boolean
    Dim oTxt 'Modify by Amy 2021/07/27 改From2.0 拿掉 As TextBox
    
    TxtValidate = False
    For Each oTxt In Text1
        If oTxt <> MsgText(601) Then
            TxtValidate = True
            Exit For
        End If
    Next
    For Each oTxt In Text2
        If oTxt <> MsgText(601) Then
            TxtValidate = True
            Exit For
        End If
    Next
    If Text2(0) <> MsgText(601) Then Call Text2_Validate(0, False)
    If cboStatus <> MsgText(601) Then
        TxtValidate = True
        If (cboStatus = "遺失" Or cboStatus = "銷毀") And Check1.Value = 0 Then Check1.Value = 1
    End If
    If cboClass <> MsgText(601) Then TxtValidate = True
    
End Function

Private Sub SetGridWidth()
   
    '設欄寬
    With grdDataList
        .FormatString = .FormatString
        For i = LBound(intWidth) To UBound(intWidth)
            .ColWidth(i) = intWidth(i)
            If intWidth(i) <> 0 Then .ColAlignment(i) = flexAlignLeftCenter
        Next i
    End With
End Sub

Private Function GetValue(strField As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(arrField)
       If UCase(arrField(jj)) = UCase(strField) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Public Sub PubShowNextData()
    Dim j As Integer
    
    With grdDataList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "V" Then
                .TextMatrix(i, 0) = ""
                grdDataList.row = i
                For j = GetValue("ISBN") To GetValue("出刊日")
                    .col = j
                    .CellBackColor = QBColor(15)
                Next j
                If fnSaveParentForm(Me) = False Then
                   Me.Enabled = True
                   Exit Sub
                End If
                Screen.MousePointer = vbHourglass
                Select Case cmdState
                    Case 0
                        If frm010034.QueryRecord(.TextMatrix(i, GetValue("編號"))) = True Then
                            frm010034.SetParent Me
                            frm010034.ToolBarSet -1
                            frm010034.Show
                            Screen.MousePointer = vbDefault
                            Me.Enabled = True
                            Exit Sub
                        End If
                    Case 1
                        If frm010035_3.QueryRecord(.TextMatrix(i, GetValue("編號"))) = True Then
                            frm010035_3.Show
                            Screen.MousePointer = vbDefault
                            Me.Enabled = True
                            Exit Sub
                        End If
                        
                End Select
            End If
        Next i
        Screen.MousePointer = vbDefault
    End With
End Sub

Private Sub ClearField()
    Dim oText 'Modify by Amy 2021/07/27 改From2.0 拿掉 As TextBox
    
    For Each oText In Text1
        oText.Text = Empty
    Next
    For Each oText In Text2
        oText.Text = Empty
    Next
    
    LblEmp = Empty
    cboClass = ""
    cboStatus = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frm010036 = Nothing
End Sub

Private Sub GrdDataList_Click()
    With grdDataList
        .Visible = False
        .col = 0
        If .row <> 0 Then
            If .Text = "V" Then
                .Text = ""
                For i = GetValue("ISBN") To GetValue("出刊日")
                    .col = i
                    .CellBackColor = QBColor(15)
                Next i
            Else
                .Text = "V"
                For i = GetValue("ISBN") To GetValue("出刊日")
                    .col = i
                   .CellBackColor = &HFFC0C0
                Next i
            End If
        End If
        .Visible = True
    End With
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    TextInverse Text1(Index)
End Sub

Private Sub Text2_GotFocus(Index As Integer)
    TextInverse Text2(Index)
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Index As Integer, Cancel As Boolean)
    If Trim(Text2(Index)) = MsgText(601) Then Exit Sub
    
    Select Case Index
        Case 0
            LblEmp = StaffQuery(Text2(0))
            If LblEmp = MsgText(601) Then
                Text2(Index).SetFocus: Cancel = True: Exit Sub
            End If
        Case 1, 2
            If CheckIsTaiwanDate(Text2(Index)) = False Then
                Text2(Index).SetFocus: Cancel = True: Exit Sub
            End If
    End Select
    
End Sub
