VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm010035_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "借閱記錄"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   1935
   ClientWidth     =   6315
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   6315
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面"
      Height          =   400
      Left            =   5340
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   10
      Width           =   950
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GrdDataList 
      Height          =   2925
      Left            =   120
      TabIndex        =   0
      Top             =   2085
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   5159
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "借閱日|借閱人|應還日|歸還日|延期次數|借閱備註"
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
      _Band(0).Cols   =   6
   End
   Begin MSForms.ComboBox CboBK 
      Height          =   300
      Index           =   1
      Left            =   1080
      TabIndex        =   23
      Top             =   1000
      Width           =   5100
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "8996;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox CboBK 
      Height          =   300
      Index           =   0
      Left            =   1080
      TabIndex        =   22
      Top             =   705
      Width           =   5100
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "8996;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   13
      Left            =   1080
      TabIndex        =   21
      Top             =   1860
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblBK(13)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   11
      Left            =   4050
      TabIndex        =   20
      Top             =   1605
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblBK(11)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   10
      Left            =   1080
      TabIndex        =   19
      Top             =   1600
      Width           =   1305
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
      Index           =   12
      Left            =   4050
      TabIndex        =   18
      Top             =   1350
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblBK(12)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   8
      Left            =   1080
      TabIndex        =   17
      Top             =   1350
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblBK(8)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   9
      Left            =   4050
      TabIndex        =   16
      Top             =   420
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblBK(9)"
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
      TabIndex        =   15
      Top             =   420
      Width           =   1305
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
      Left            =   4050
      TabIndex        =   14
      Top             =   120
      Width           =   1305
      VariousPropertyBits=   27
      Caption         =   "LblBK(2)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblBK 
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   13
      Top             =   120
      Width           =   960
      VariousPropertyBits=   27
      Caption         =   "LblBK(1)"
      Size            =   "11721;503"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "出刊日期："
      Height          =   255
      Index           =   14
      Left            =   120
      TabIndex        =   12
      Top             =   1860
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "保管人員："
      Height          =   255
      Index           =   11
      Left            =   120
      TabIndex        =   11
      Top             =   1600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "保管單位："
      Height          =   255
      Index           =   12
      Left            =   3075
      TabIndex        =   10
      Top             =   1605
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "狀　　態："
      Height          =   255
      Index           =   13
      Left            =   3075
      TabIndex        =   9
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "譯　　者："
      Height          =   255
      Index           =   10
      Left            =   120
      TabIndex        =   8
      Top             =   1350
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "圖書編號："
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "類　　別："
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   420
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "書名 (中)："
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   700
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "ＩＳＢＮ："
      Height          =   255
      Index           =   1
      Left            =   3075
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "作者 (中)："
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1000
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "上架日期："
      Height          =   255
      Index           =   7
      Left            =   3090
      TabIndex        =   2
      Top             =   420
      Width           =   1095
   End
End
Attribute VB_Name = "frm010035_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2021/07/27 Form2.0已修改 LblBK/GrdDataList
'2016/10/03 Create by Amy
Option Explicit

Dim i As Integer
Dim arrField, intWidth
Public strPreFormName As String '上層FormName

Private Sub cmdBack_Click()
    If strPreFormName = "" Then
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Else
        frm010034.Show
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    ReDim arrField(6)
    ReDim intWidth(6)
    arrField = Array("借閱日", "借閱人", "應還日", "歸還日", "延期次數", "借閱備註", "LR01", "LR03")
    intWidth = Array(800, 800, 800, 800, 800, 1800, 0, 0)
    
    ClearField
    MoveFormToCenter Me
End Sub

Public Function QueryRecord(ByVal stBK01 As String) As Boolean
    Dim oLbl, idx As Integer  'Modify by Amy 2021/07/27 原:oLbl As LABEL
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strTmp As String
    
    QueryRecord = False
        
    strQ = "Select b.*,ST02 From BooksData b,Staff Where BK01='" & stBK01 & "' And BK10=ST01(+) "
    RsQ.CursorLocation = adUseClient
    RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
    If RsQ.RecordCount > 0 Then
        ClearField
        With RsQ
            For Each oLbl In LblBK
                 idx = oLbl.Index
                 If Not IsNull(RsQ.Fields("BK" & Format(idx, "00"))) Then
                    Select Case idx
                        Case 3
                            strTmp = GetBK03(RsQ.Fields("BK" & Format(idx, "00")))
                        Case 9, 13
                            strTmp = Format(TAIWANDATE(RsQ.Fields("BK" & Format(idx, "00"))), "###/##/##")
                        Case 10
                            strTmp = RsQ.Fields("BK" & Format(idx, "00")) & " " & RsQ.Fields("ST02")
                        Case Else
                            strTmp = RsQ.Fields("BK" & Format(idx, "00"))
                    End Select
                    oLbl.Caption = strTmp 'Modify by Amy 2021/07/27 +.Caption
                 End If
            Next
            '書名
            If Not IsNull(RsQ.Fields("BK04")) Then CboBK(0).AddItem "中 : " & RsQ.Fields("BK04")
            If Not IsNull(RsQ.Fields("BK05")) Then CboBK(0).AddItem "英 : " & RsQ.Fields("BK05")
            CboBK(0).ListIndex = 0
            '作者
            If Not IsNull(RsQ.Fields("BK06")) Then CboBK(1).AddItem "中 : " & RsQ.Fields("BK06")
            If Not IsNull(RsQ.Fields("BK07")) Then CboBK(1).AddItem "英 : " & RsQ.Fields("BK07")
            CboBK(1).ListIndex = 0
        End With
        strQ = "Select sqldatet(L.LR04) as 借閱,ST02 as 借閱人,sqldatet(L.LR05) as 應還日, sqldatet(R.LR09) as 歸還日,Nvl(CFreq,0) as 延期次數,L.LR07 as 備註,O.LR01,O.LR03  From LoanRecord O," & _
                "(Select LR01,LR02,LR04,LR05,LR07,LR08 From LoanRecord Where LR03='" & stBK01 & "' And LR02='1') L," & _
                "(Select LR01,LR09 From LoanRecord Where LR03='" & stBK01 & "' And LR02='X') R," & _
                "(Select LR01,Count(*) as CFreq From LoanRecord Where LR03='" & stBK01 & "' And LR02>'1' And LR02<'X' Group by LR01) C,Staff " & _
                "Where O.LR01=L.LR01(+) And O.LR01=R.LR01(+) And O. LR01=C.LR01(+) And L.LR08=ST01(+) And O.LR03='" & stBK01 & "' " & _
                "Group by O.LR01,O.LR03,L.LR04, R.LR09,L.LR05, ST02,CFreq,L.LR07 "
        If RsQ.State = adStateOpen Then RsQ.Close
        RsQ.CursorLocation = adUseClient
        RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
        grdDataList.Clear
        grdDataList.Rows = 2
        Set grdDataList.Recordset = RsQ
        SetGridWidth
        QueryRecord = True
    End If
    RsQ.Close
End Function

Private Sub ClearField()
    Dim lbl, cbo 'Modify by Amy 2021/07/27 原:Lbl As LABEL, cbo As ComboBox
    
    For Each lbl In LblBK
        lbl.Caption = Empty
    Next
    For Each cbo In CboBK
       cbo.Clear
    Next
    
End Sub

Private Sub SetGridWidth()
    '設欄寬
    With grdDataList
        .FormatString = .FormatString
        For i = LBound(intWidth) To UBound(intWidth)
            .ColWidth(i) = intWidth(i)
            If intWidth(i) <> 0 Then
                .ColAlignment(i) = flexAlignLeftCenter
                If i = GetValue("延期次數") Then .ColAlignment(i) = flexAlignRightCenter
            End If
        Next i
    End With
End Sub

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

Private Sub Form_Unload(Cancel As Integer)
    strPreFormName = ""
    Set frm010035_3 = Nothing
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

