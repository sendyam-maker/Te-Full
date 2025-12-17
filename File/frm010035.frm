VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010035 
   BorderStyle     =   1  '單線固定
   Caption         =   "圖書借閱查詢"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   7560
   Begin VB.CheckBox Check1 
      Caption         =   "含已遺失或銷毀"
      Height          =   300
      Left            =   3600
      TabIndex        =   32
      Top             =   1800
      Width           =   1800
   End
   Begin VB.CommandButton cmdLoanRecord 
      Caption         =   "個人記錄"
      Height          =   450
      Left            =   100
      Style           =   1  '圖片外觀
      TabIndex        =   14
      Top             =   45
      Width           =   950
   End
   Begin VB.ComboBox cboClass 
      Height          =   300
      ItemData        =   "frm010035.frx":0000
      Left            =   1080
      List            =   "frm010035.frx":0016
      TabIndex        =   5
      Text            =   "cboClass"
      Top             =   1800
      Width           =   1680
   End
   Begin VB.CommandButton cmdNewBooks 
      Caption         =   "三個月內新書"
      Height          =   450
      Left            =   5655
      Style           =   1  '圖片外觀
      TabIndex        =   11
      Top             =   45
      Width           =   950
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "借閱記錄"
      Height          =   400
      Index           =   2
      Left            =   6450
      Style           =   1  '圖片外觀
      TabIndex        =   22
      Top             =   2560
      Width           =   950
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "申請借閱"
      Height          =   400
      Index           =   0
      Left            =   4500
      Style           =   1  '圖片外觀
      TabIndex        =   18
      Top             =   2560
      Width           =   950
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "書籍資料"
      Height          =   400
      Index           =   1
      Left            =   5480
      Style           =   1  '圖片外觀
      TabIndex        =   20
      Top             =   2560
      Width           =   950
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   450
      Left            =   6660
      Style           =   1  '圖片外觀
      TabIndex        =   12
      Top             =   45
      Width           =   800
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "尋找"
      Default         =   -1  'True
      Height          =   450
      Left            =   4800
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   45
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2450
      Left            =   120
      TabIndex        =   13
      Top             =   3100
      Width           =   7245
      _ExtentX        =   12779
      _ExtentY        =   4313
      _Version        =   393216
      Cols            =   13
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "V|編號|書名|ISBN|作者|譯者|類別|保管人|狀態|借閱人|借閱/延期日|上架日|出刊日"
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
      Height          =   270
      Index           =   4
      Left            =   1080
      TabIndex        =   4
      Top             =   1500
      Width           =   5300
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   8
      Left            =   5880
      TabIndex        =   9
      Top             =   2100
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   7
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   7
      Left            =   4560
      TabIndex        =   8
      Top             =   2100
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   7
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   6
      Left            =   2400
      TabIndex        =   7
      Top             =   2100
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   7
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   4
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   2
      Left            =   1080
      TabIndex        =   2
      Top             =   900
      Width           =   5300
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   1
      Left            =   4380
      TabIndex        =   1
      Top             =   600
      Width           =   2000
      VariousPropertyBits=   679493659
      MaxLength       =   20
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   3
      Left            =   1080
      TabIndex        =   3
      Top             =   1200
      Width           =   5300
      VariousPropertyBits=   679493659
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox Text1 
      Height          =   270
      Index           =   5
      Left            =   1080
      TabIndex        =   6
      Top             =   2100
      Width           =   855
      VariousPropertyBits=   679493659
      MaxLength       =   7
      Size            =   "8555;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "譯　　者："
      Height          =   255
      Index           =   12
      Left            =   120
      TabIndex        =   31
      Top             =   1500
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(模糊比對)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   11
      Left            =   6480
      TabIndex        =   30
      Top             =   1555
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(模糊比對)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   10
      Left            =   6480
      TabIndex        =   29
      Top             =   1270
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "(模糊比對)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Index           =   9
      Left            =   6480
      TabIndex        =   28
      Top             =   960
      Width           =   840
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7320
      Y1              =   2520
      Y2              =   2520
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
      Left            =   5580
      TabIndex        =   27
      Top             =   2145
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "出刊日期："
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   26
      Top             =   2100
      Width           =   1095
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
      Index           =   5
      Left            =   2100
      TabIndex        =   25
      Top             =   2145
      Width           =   300
   End
   Begin VB.Label Label1 
      Caption         =   "圖書編號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   24
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "類　　別："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "書　　名："
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ＩＳＢＮ："
      Height          =   255
      Index           =   1
      Left            =   3465
      TabIndex        =   19
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "(流水號)"
      Height          =   255
      Left            =   1965
      TabIndex        =   17
      Top             =   600
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   "作　　者："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   16
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "上架日期："
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   15
      Top             =   2100
      Width           =   1095
   End
End
Attribute VB_Name = "frm010035"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/27 Form2.0已修改 Text1/GrdDataList
'2016/10/03 Create by Amy
Option Explicit

Public cmdState As Integer
Public bolLoanRecordApply As Boolean '保管人是否需簽核
Dim bolNewBooks As Boolean '近期新書
Dim i As Integer
Dim arrField, intWidth

Private Sub cboClass_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 '設只能選
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Public Sub cmdLoanRecord_Click()
    If fnSaveParentForm(Me) = False Then
        Me.Enabled = True
        Exit Sub
    End If
    SetGridNoV
    If bolLoanRecordApply = True Then
        bolLoanRecordApply = False
    End If
    frm010035_2.Option1(2).Value = True
    Call frm010035_2.cmdSearch_Click
    frm010035_2.Show
End Sub

Private Sub cmdNewBooks_Click()
    bolNewBooks = True
    Call cmdSearch_Click
    bolNewBooks = False
End Sub

Private Sub cmdok_Click(Index As Integer)
    cmdState = Index
    PubShowNextData
End Sub

Private Sub cmdSearch_Click()
    Dim oTxt 'Modify by Amy 2021/07/29 原:As TextBox
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String
    
    If bolNewBooks = True Then
        strQ = strQ & " And BK09>=" & Val(DBDATE(DateAdd("m", -3, Format(strSrvDate(1), "####/##/##")))) & _
                " And BK09<=" & Val(strSrvDate(1))
    Else
        If TxtValidate = False Then MsgBox "請輸入查詢條件！", vbInformation: Exit Sub
        
        For Each oTxt In Text1
            If oTxt <> MsgText(601) Then
                Select Case oTxt.Index
                    Case 0
                        strQ = strQ & " And BK01='" & oTxt & "' "
                    Case 1
                         strQ = strQ & " And BK02='" & oTxt & "' "
                    Case 2
                        strQ = strQ & " And (Upper(BK04) like '%" & UCase(oTxt) & "%' Or Upper(BK05) like '%" & UCase(oTxt) & "%')"
                    Case 3
                         strQ = strQ & " And (Upper(BK06) like '%" & UCase(oTxt) & "%' Or Upper(BK07) like '%" & UCase(oTxt) & "%')"
                    Case 4
                         strQ = strQ & " And BK08 like '%" & UCase(oTxt) & "%'"
                    Case 5
                         strQ = strQ & " And BK09>=" & Val(oTxt) + 19110000
                    Case 6
                         strQ = strQ & " And BK09<=" & Val(oTxt) + 19110000
                    Case 7
                         strQ = strQ & " And BK13>=" & Val(oTxt) + 19110000
                    Case 8
                         strQ = strQ & " And BK13<=" & Val(oTxt) + 19110000
                End Select
            End If
        Next
        If cboClass <> MsgText(601) Then strQ = strQ & " And BK03=" & Mid(cboClass, 1, InStr(cboClass, ".") - 1)
    End If
    
    strQ = "Select '' as V,BK01 as 編號,Decode(BK04,null,BK05,Decode(BK05,null,BK04,BK04||'('||BK05||')')) as 書名,Nvl(BK02,'') as ISBN," & _
                "Decode(BK06,null,BK07,Decode(BK07,null,BK06,BK06||'('||BK07||')')) as 作者,Nvl(BK08,'') as 譯者," & _
                "Decode(BK03,'1','專利','2','商標','3','法律','4','電腦','5','其他') as 類別,k.ST02 as 保管人," & _
                "Decode(LR06,null,Decode(LR02, '1','借閱申請中', 'X','可借閱', 'Y','遺失', 'Z','銷毀', '延期申請中'), Decode(LR02, 'X','可借閱', 'Y','遺失', 'Z','銷毀', '借閱中')) as  狀態," & _
                "Decode(LR02,'X','',b.ST02) as 借閱人,Decode(LR02,'Z','',Decode(LR06,null,'',sqldatet(LR04))) as 借閱日," & _
                "sqldatet(BK09) as 上架日,Decode(BK13,null,'',sqldatet(BK13)) as 出刊日, " & _
                "Nvl(LR06,'') as LR06,BK10,LR01,LR02,LR08 From (Select * From BooksData Where " & Mid(strQ, 5) & ")," & _
                "(Select * From LoanRecord a Where LR01||LR02=(Select Max(LR01||LR02) From LoanRecord  Where a.LR03=LR03 ))" & _
                ",Staff k,Staff b Where BK01=LR03(+) And BK10=k.ST01(+) And LR08=b.ST01(+) " & _
                IIf(Check1.Value = 0, " And LR02<>'Y' And LR02<>'Z' ", "")
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    grdDataList.Clear
    grdDataList.FixedCols = 0
    If RsQ.RecordCount > 0 Then
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
    SetGridWidth
    grdDataList.FixedCols = 3 '固定前3欄
    cboClass = ""
    cmdState = -1
End Sub

Private Function TxtValidate() As Boolean
    Dim oTxt 'Modify by Amy 2021/07/29 原:As TextBox
    
    TxtValidate = False
    For Each oTxt In Text1
        If oTxt <> MsgText(601) Then
            TxtValidate = True
            Exit For
        End If
    Next
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
                .row = i
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
                    Case 0 '申請借閱
                        If .TextMatrix(i, GetValue("狀態")) = "可借閱" Then
                            If frm010035_1.QueryRecord(.TextMatrix(i, GetValue("編號"))) = True Then
                                frm010035_1.strPreRow = i
                                frm010035_1.Show
                                Screen.MousePointer = vbDefault
                                Me.Enabled = True
                                Exit Sub
                            End If
                        Else
                            Me.Show
                            MsgBox "此書目前狀態為「" & .TextMatrix(i, GetValue("狀態")) & "」,故不可借閱！", vbInformation
                        End If
                    Case 1 '書籍資料
                        If frm010034.QueryRecord(.TextMatrix(i, GetValue("編號"))) = True Then
                            frm010034.SetParent Me
                            frm010034.ToolBarSet -1
                            frm010034.Show
                            Screen.MousePointer = vbDefault
                            Me.Enabled = True
                            Exit Sub
                        End If
                    Case 2 '借閱記錄
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

Private Sub Form_Unload(Cancel As Integer)
    bolLoanRecordApply = False
    Set frm010035 = Nothing
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
    
    Select Case Index
        Case 0, 1, 5, 6, 7, 8
            CloseIme
    End Select
End Sub

Private Sub SetGridNoV()
    Dim j As Integer
    
    With grdDataList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 0) = "V" Then
                .TextMatrix(i, 0) = ""
                .row = i
                For j = GetValue("ISBN") To GetValue("出刊日")
                    .col = j
                    .CellBackColor = QBColor(15)
                Next j
            End If
        Next i
    End With
End Sub
