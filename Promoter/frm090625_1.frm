VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm090625_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "工程師每週完稿明細"
   ClientHeight    =   5715
   ClientLeft      =   1650
   ClientTop       =   1515
   ClientWidth     =   9315
   ControlBox      =   0   'False
   FillColor       =   &H80000005&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   9315
   Visible         =   0   'False
   Begin VB.CommandButton cmdok 
      Caption         =   "開啟Word(&W)"
      Height          =   400
      Index           =   2
      Left            =   6780
      TabIndex        =   5
      Top             =   0
      Width           =   1260
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Index           =   1
      Left            =   5610
      TabIndex        =   3
      Top             =   0
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   0
      Left            =   8100
      TabIndex        =   0
      Top             =   0
      Width           =   1140
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd 
      Height          =   5130
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   450
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   9049
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      HighLight       =   2
      AllowUserResizing=   3
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
      _Band(0).Cols   =   2
   End
   Begin VB.Label lblMonth 
      Caption         =   "lblMonth"
      Height          =   180
      Left            =   1140
      TabIndex        =   2
      Top             =   60
      Width           =   1065
   End
   Begin VB.Label Label1 
      Caption         =   "完稿月份： "
      Height          =   180
      Index           =   35
      Left            =   150
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
End
Attribute VB_Name = "frm090625_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; grd(0)改字型=新細明體-ExtB ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit
Dim m_dblRow As Double
Dim m_dblCol As Double
Dim PLeft(0 To 7) As Integer
Dim m_intPage As Integer
Dim m_iPrint As Integer

Private Sub cmdOK_Click(Index As Integer)
Dim ii As Integer
    
    Select Case Index
    Case 0 '回前畫面
        Unload Me
    Case 1 '列印
        Screen.MousePointer = vbHourglass
        PrintData
        Screen.MousePointer = vbDefault
    Case 2 '產生Word
        Screen.MousePointer = vbHourglass
        Me.grd(0).MousePointer = vbHourglass
        OpenWord
        Me.grd(0).MousePointer = vbDefault
        Screen.MousePointer = vbDefault
    Case Else
    End Select
End Sub

Private Sub Form_Activate()
    If Me.grd(0).Rows < 2 Then
        Unload Me
    ElseIf Me.grd(0).Rows = 2 Then
        If Me.grd(0).TextMatrix(1, 0) = "" Then
            Unload Me
        End If
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = vbHourglass
    MoveFormToCenter Me
    Me.lblMonth.Caption = frm090625.txt1(0).Text
    SetGrd
    Process
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frm090625.Show
    Set frm090625_1 = Nothing
End Sub

Private Sub SetGrd()
Dim ii As Integer

        With Me.grd(ii)
            .Visible = False
            .Cols = 8
            .row = 0
            .col = 0: .ColWidth(0) = 900
            .CellAlignment = flexAlignCenterCenter
            .Text = "員工姓名"
            .col = 1: .ColWidth(1) = 700
            .CellAlignment = flexAlignCenterCenter
            .Text = "週次"
            .col = 2: .ColWidth(2) = 1200
            .CellAlignment = flexAlignCenterCenter
            .Text = "I"
            .col = 3: .ColWidth(3) = 1200
            .CellAlignment = flexAlignCenterCenter
            .Text = "III"
            .col = 4: .ColWidth(4) = 1200
            .CellAlignment = flexAlignCenterCenter
            .Text = "RE"
            .col = 5: .ColWidth(5) = 1200
            .CellAlignment = flexAlignCenterCenter
            .Text = "OT"
            .col = 6: .ColWidth(6) = 1200
            .CellAlignment = flexAlignCenterCenter
            .Text = "PI"
            .col = 7: .ColWidth(7) = 1200
            .CellAlignment = flexAlignCenterCenter
            .Text = "PIII"
            
            .MergeCells = flexMergeRestrictRows
'            .MergeRow(0) = True
            .MergeCol(0) = True
'            .MergeRow(1) = True
            .MergeCol(1) = True
'            .ColWidth(0) = 800
'            .ColWidth(1) = 1200
            .Visible = True
        End With
'    '預設目前在第一筆的位置
'    With Me.grd(0)
'        .Row = 0
'        .col = 2
'    End With
End Sub

Private Sub Process()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSQLA1 As String
Dim strSQLA2 As String
Dim strSQLA3 As String
Dim strSQLA4 As String
Dim strCPKind As String
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim ii As Integer

StrSQLa = "Delete From R090625 Where ID='" & strUserNum & "' "
cnnConnection.Execute StrSQLa
StrSQLa = "Delete From R090625_1 Where ID='" & strUserNum & "' "
cnnConnection.Execute StrSQLa
With frm090625
    StrSQLa = "": strSQLA1 = "": strSQLA2 = "": strSQLA3 = "": strSQLA4 = ""
'    strSQLA = strSQLA & " And ST16='CFP' And CP26 Is Null And CP01 In ('P','CFP') "
    StrSQLa = StrSQLa & " And ST16 In ('P','CFP') And CP26 Is Null And CP01 In ('P','CFP') "
    '所別
    If .txt1(9).Text <> "" Then
        StrSQLa = StrSQLa & " And ST06>='" & .txt1(9).Text & "' "
    End If
    If .txt1(10).Text <> "" Then
        StrSQLa = StrSQLa & " And ST06<='" & .txt1(10).Text & "' "
    End If
    If .txt1(9).Text <> "" Or .txt1(10).Text <> "" Then
         pub_QL05 = pub_QL05 & ";" & Left(frm090625.Label1(10), 3) & frm090625.txt1(9) & "-" & frm090625.txt1(10) & "(1:北所 2:中所 3:南所 4:高所 5:其他)" 'Add By Sindy 2010/12/20
    End If
    '員工編號
    If .txt1(11).Text <> "" Then
        StrSQLa = StrSQLa & " And ST01='" & .txt1(11).Text & "' "
        pub_QL05 = pub_QL05 & ";" & frm090625.Label1(11) & frm090625.txt1(11) & frm090625.Label1(12) 'Add By Sindy 2010/12/20
    End If
    '第一週
    strSQLA1 = strSQLA1 & " And EP09>=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(1).Text, "00")) & " And EP09<=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(2).Text, "00")) & " "
    '第二週
    strSQLA2 = strSQLA2 & " And EP09>=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(3).Text, "00")) & " And EP09<=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(4).Text, "00")) & " "
    '第三週
    strSQLA3 = strSQLA3 & " And EP09>=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(5).Text, "00")) & " And EP09<=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(6).Text, "00")) & " "
    '第四週
    strSQLA4 = strSQLA4 & " And EP09>=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(7).Text, "00")) & " And EP09<=" & ChangeTStringToWString(.txt1(0).Text & Format(.txt1(8).Text, "00")) & " "
    pub_QL05 = pub_QL05 & ";" & Left(frm090625.Label1(2), 5) & frm090625.txt1(0) 'Add By Sindy 2010/12/20
    pub_QL05 = pub_QL05 & ";" & frm090625.Label1(3) & frm090625.txt1(1) & "-" & frm090625.txt1(2) 'Add By Sindy 2010/12/20
    pub_QL05 = pub_QL05 & ";" & frm090625.Label1(0) & frm090625.txt1(3) & "-" & frm090625.txt1(4) 'Add By Sindy 2010/12/20
    pub_QL05 = pub_QL05 & ";" & frm090625.Label1(1) & frm090625.txt1(5) & "-" & frm090625.txt1(6) 'Add By Sindy 2010/12/20
    pub_QL05 = pub_QL05 & ";" & frm090625.Label1(4) & frm090625.txt1(7) & "-" & frm090625.txt1(8) 'Add By Sindy 2010/12/20
    
    strSql = "Select '1', EP05, CP01, CP02, CP03, CP04, CP10 From Engineerprogress, Caseprogress, Staff Where EP02=CP09 And EP05=ST01 " & strSQLA1 & StrSQLa
    strSql = strSql & " Union Select '2', EP05, CP01, CP02, CP03, CP04, CP10 From Engineerprogress, Caseprogress, Staff Where EP02=CP09 And EP05=ST01 " & strSQLA2 & StrSQLa
    strSql = strSql & " Union Select '3', EP05, CP01, CP02, CP03, CP04, CP10 From Engineerprogress, Caseprogress, Staff Where EP02=CP09 And EP05=ST01 " & strSQLA3 & StrSQLa
    strSql = strSql & " Union Select '4', EP05, CP01, CP02, CP03, CP04, CP10 From Engineerprogress, Caseprogress, Staff Where EP02=CP09 And EP05=ST01 " & strSQLA4 & StrSQLa
    rsA.CursorLocation = adUseClient
    rsA.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            strCPKind = ""
            '下列設定的案件性質代碼可能系統類別(P, CFP)並未設定, 但為避免往後又設定歸屬後本程式並未計入, 因此仍設定進去
            Select Case "" & rsA.Fields(2).Value
            Case "P"
                Select Case "" & rsA.Fields(6).Value
                Case "215", "606", "997", "998"
                    strCPKind = "4"
                Case "101", "102", "104", "107", "108", "109", "110", "111", "112", "116", "117", "119", "120", "201", "205", "301", "302", "304", "305", "307", "501", "502", "503", "504", "507", "508", "801", "802", "803", "804", "906"
                    strCPKind = "5"
                Case "103", "105", "106", "202", "203", "204", "206", "207", "209", "215", "303", "306", "401", "402", "403", "404", "405", "406", "407", "408", "409", "410", "411", "412", "413", "414", "415", "416", "417", "418", "419", "420", "421", "425", "426", "505", "506", "601", "602", "603", "604", "605", "606", "608", _
                        "701", "702", "703", "704", "705", "706", "707", "708", "901", "902", "903", "904", "905", "907", "908", "909", "910", "911", "912", "913", "915", "916", "917", "938", "939", "997", "998"
                    strCPKind = "6"
                End Select
            Case "CFP"
                Select Case "" & rsA.Fields(6).Value
                'Modified by Morgan 2017/11/6 +122 --王副總
                Case "101", "102", "113", "115", "118", "122", "301", "302", "306", "307", "501", "502", "801", "802", "803", "804", "805", "906"
                    strCPKind = "1"
                Case "103", "105", "114", "303", "305"
                    strCPKind = "2"
                Case "107", "424"
                    strCPKind = "3"
                Case "106", "108", "116", "120", "201", "202", "203", "204", "206", "207", "208", "214", "215", "216", "217", "401", "402", "403", "404", "405", "407", "408", "413", "414", "416", "421", "422", "423", "427", "503", "601", "604", "605", "606", "607", "701", "702", "704", "705", "902", "903", "904", "907", "909", "910", "911", "913", "914", "917", "938", "939", "997", "998", "999"
                    strCPKind = "4"
                End Select
            End Select
            If strCPKind <> "" Then
                StrSQLa = "Insert Into R090625 Values('" & rsA.Fields(1).Value & "','" & rsA.Fields(0).Value & "','" & strCPKind & "','" & rsA.Fields(2).Value & "','" & rsA.Fields(3).Value & "','" & rsA.Fields(4).Value & "','" & rsA.Fields(5).Value & "','" & strUserNum & "') "
                cnnConnection.Execute StrSQLa
            End If
            rsA.MoveNext
        Wend
        StrSQLa = "Insert Into R090625(R09062501, R09062502, R09062503, R09062505, ID) Select R09062501, '合　計', R09062503, To_Char(Count(*)), '" & strUserNum & "' From R090625 Where ID='" & strUserNum & "' Group By R09062501, R09062503 "
        cnnConnection.Execute StrSQLa
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        StrSQLa = "Select R09062501, R09062502, R09062503 From R090625 Where ID='" & strUserNum & "' Group By R09062501, R09062502, R09062503 Order By 1, 2, 3 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        If rsA.RecordCount > 0 Then
            While Not rsA.EOF
                StrSqlB = "Select * From R090625 Where R09062501='" & rsA.Fields(0).Value & "' And R09062502='" & rsA.Fields(1).Value & "' And R09062503='" & rsA.Fields(2).Value & "' And ID='" & strUserNum & "' "
                rsB.CursorLocation = adUseClient
                rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
                If rsB.RecordCount > 0 Then
                    ii = 0
                    While Not rsB.EOF
                        ii = ii + 1
                        '若有重覆
                        If ChkDataDuplicate(rsB.Fields(0).Value, rsB.Fields(1).Value, "" & ii) = True Then
                            Select Case "" & rsB.Fields(2).Value
                            Case "1"
                                StrSqlB = "Update R090625_1 Set R090625_104='" & rsB.Fields(3).Value & "', R090625_105='" & rsB.Fields(4).Value & "', R090625_106='" & rsB.Fields(5).Value & "', R090625_107='" & rsB.Fields(6).Value & "' Where R090625_101='" & rsB.Fields(0).Value & "' And R090625_102='" & rsB.Fields(1).Value & "' And R090625_103=" & ii & " And ID='" & strUserNum & "' "
                            Case "2"
                                StrSqlB = "Update R090625_1 Set R090625_108='" & rsB.Fields(3).Value & "', R090625_109='" & rsB.Fields(4).Value & "', R090625_110='" & rsB.Fields(5).Value & "', R090625_111='" & rsB.Fields(6).Value & "' Where R090625_101='" & rsB.Fields(0).Value & "' And R090625_102='" & rsB.Fields(1).Value & "' And R090625_103=" & ii & " And ID='" & strUserNum & "' "
                            Case "3"
                                StrSqlB = "Update R090625_1 Set R090625_112='" & rsB.Fields(3).Value & "', R090625_113='" & rsB.Fields(4).Value & "', R090625_114='" & rsB.Fields(5).Value & "', R090625_115='" & rsB.Fields(6).Value & "' Where R090625_101='" & rsB.Fields(0).Value & "' And R090625_102='" & rsB.Fields(1).Value & "' And R090625_103=" & ii & " And ID='" & strUserNum & "' "
                            Case "4"
                                StrSqlB = "Update R090625_1 Set R090625_116='" & rsB.Fields(3).Value & "', R090625_117='" & rsB.Fields(4).Value & "', R090625_118='" & rsB.Fields(5).Value & "', R090625_119='" & rsB.Fields(6).Value & "' Where R090625_101='" & rsB.Fields(0).Value & "' And R090625_102='" & rsB.Fields(1).Value & "' And R090625_103=" & ii & " And ID='" & strUserNum & "' "
                            Case "5"
                                StrSqlB = "Update R090625_1 Set R090625_120='" & rsB.Fields(3).Value & "', R090625_121='" & rsB.Fields(4).Value & "', R090625_122='" & rsB.Fields(5).Value & "', R090625_123='" & rsB.Fields(6).Value & "' Where R090625_101='" & rsB.Fields(0).Value & "' And R090625_102='" & rsB.Fields(1).Value & "' And R090625_103=" & ii & " And ID='" & strUserNum & "' "
                            Case "6"
                                StrSqlB = "Update R090625_1 Set R090625_124='" & rsB.Fields(3).Value & "', R090625_125='" & rsB.Fields(4).Value & "', R090625_126='" & rsB.Fields(5).Value & "', R090625_127='" & rsB.Fields(6).Value & "' Where R090625_101='" & rsB.Fields(0).Value & "' And R090625_102='" & rsB.Fields(1).Value & "' And R090625_103=" & ii & " And ID='" & strUserNum & "' "
                            End Select
                            cnnConnection.Execute StrSqlB
                        '若不重覆
                        Else
                            Select Case "" & rsB.Fields(2).Value
                            Case "1"
                                StrSqlB = "Insert Into R090625_1(R090625_101, R090625_102, R090625_103, R090625_104, R090625_105, R090625_106, R090625_107, ID) Values('" & rsB.Fields(0).Value & "','" & rsB.Fields(1).Value & "'," & ii & ",'" & rsB.Fields(3).Value & "','" & rsB.Fields(4).Value & "','" & rsB.Fields(5).Value & "','" & rsB.Fields(6).Value & "','" & strUserNum & "') "
                            Case "2"
                                StrSqlB = "Insert Into R090625_1(R090625_101, R090625_102, R090625_103, R090625_108, R090625_109, R090625_110, R090625_111, ID) Values('" & rsB.Fields(0).Value & "','" & rsB.Fields(1).Value & "'," & ii & ",'" & rsB.Fields(3).Value & "','" & rsB.Fields(4).Value & "','" & rsB.Fields(5).Value & "','" & rsB.Fields(6).Value & "','" & strUserNum & "') "
                            Case "3"
                                StrSqlB = "Insert Into R090625_1(R090625_101, R090625_102, R090625_103, R090625_112, R090625_113, R090625_114, R090625_115, ID) Values('" & rsB.Fields(0).Value & "','" & rsB.Fields(1).Value & "'," & ii & ",'" & rsB.Fields(3).Value & "','" & rsB.Fields(4).Value & "','" & rsB.Fields(5).Value & "','" & rsB.Fields(6).Value & "','" & strUserNum & "') "
                            Case "4"
                                StrSqlB = "Insert Into R090625_1(R090625_101, R090625_102, R090625_103, R090625_116, R090625_117, R090625_118, R090625_119, ID) Values('" & rsB.Fields(0).Value & "','" & rsB.Fields(1).Value & "'," & ii & ",'" & rsB.Fields(3).Value & "','" & rsB.Fields(4).Value & "','" & rsB.Fields(5).Value & "','" & rsB.Fields(6).Value & "','" & strUserNum & "') "
                            Case "5"
                                StrSqlB = "Insert Into R090625_1(R090625_101, R090625_102, R090625_103, R090625_120, R090625_121, R090625_122, R090625_123, ID) Values('" & rsB.Fields(0).Value & "','" & rsB.Fields(1).Value & "'," & ii & ",'" & rsB.Fields(3).Value & "','" & rsB.Fields(4).Value & "','" & rsB.Fields(5).Value & "','" & rsB.Fields(6).Value & "','" & strUserNum & "') "
                            Case "6"
                                StrSqlB = "Insert Into R090625_1(R090625_101, R090625_102, R090625_103, R090625_124, R090625_125, R090625_126, R090625_127, ID) Values('" & rsB.Fields(0).Value & "','" & rsB.Fields(1).Value & "'," & ii & ",'" & rsB.Fields(3).Value & "','" & rsB.Fields(4).Value & "','" & rsB.Fields(5).Value & "','" & rsB.Fields(6).Value & "','" & strUserNum & "') "
                            End Select
                            cnnConnection.Execute StrSqlB
                        End If
                        rsB.MoveNext
                    Wend
                End If
                If rsB.State <> adStateClosed Then rsB.Close
                Set rsB = Nothing
                rsA.MoveNext
            Wend
        End If
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
'        strSQLA = "Select ST02, Decode(R090625_102, '1', '第一週', '2', '第二週', '3', '第三週', '4', '第四週', R090625_102), Decode(R090625_104, Null, R090625_105, R090625_104||'-'||R090625_105||'-'||R090625_106||'-'||R090625_107) " & _
'                        ", Decode(R090625_108, Null, R090625_109, R090625_108||'-'||R090625_109||'-'||R090625_110||'-'||R090625_111) " & _
'                        ", Decode(R090625_112, Null, R090625_113, R090625_112||'-'||R090625_113||'-'||R090625_114||'-'||R090625_115) " & _
'                        ", Decode(R090625_116, Null, R090625_117, R090625_116||'-'||R090625_117||'-'||R090625_118||'-'||R090625_119) " & _
'                        ", Decode(R090625_120, Null, R090625_121, R090625_120||'-'||R090625_121||'-'||R090625_122||'-'||R090625_123) " & _
'                        ", Decode(R090625_124, Null, R090625_125, R090625_124||'-'||R090625_125||'-'||R090625_126||'-'||R090625_127) " & _
'                        " From R090625_1, Staff Where R090625_101=ST01 And ID='" & strUserNum & "' Order By ST06, R090625_101, R090625_102, R090625_103 "
        StrSQLa = "Select ST02, Decode(R090625_102, '1', '第一週', '2', '第二週', '3', '第三週', '4', '第四週', R090625_102), Decode(R090625_104, Null, R090625_105, Replace(R090625_105||'-'||R090625_106||'-'||R090625_107,'-0-00','')) " & _
                        ", Decode(R090625_108, Null, R090625_109, Replace(R090625_109||'-'||R090625_110||'-'||R090625_111,'-0-00','')) " & _
                        ", Decode(R090625_112, Null, R090625_113, Replace(R090625_113||'-'||R090625_114||'-'||R090625_115,'-0-00','')) " & _
                        ", Decode(R090625_116, Null, R090625_117, Replace(R090625_117||'-'||R090625_118||'-'||R090625_119,'-0-00','')) " & _
                        ", Decode(R090625_120, Null, R090625_121, Replace(R090625_121||'-'||R090625_122||'-'||R090625_123,'-0-00','')) " & _
                        ", Decode(R090625_124, Null, R090625_125, Replace(R090625_125||'-'||R090625_126||'-'||R090625_127,'-0-00','')) " & _
                        " From R090625_1, Staff Where R090625_101=ST01 And ID='" & strUserNum & "' Order By ST06, R090625_101, R090625_102, R090625_103 "
        rsA.CursorLocation = adUseClient
        rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
        InsertQueryLog (rsA.RecordCount) 'Add By Sindy 2010/12/20
        Set Me.grd(0).Recordset = rsA
        SetGrd
'        Me.grd(0).FixedCols = 2
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/20
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
        ShowNoData
    End If
End With
End Sub

Private Function ChkDataDuplicate(strF1 As String, strF2 As String, strF3 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

ChkDataDuplicate = False
StrSQLa = "Select * From R090625_1 Where R090625_101='" & strF1 & "' And R090625_102='" & strF2 & "' And r090625_103=" & Val(strF3) & " And ID='" & strUserNum & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    ChkDataDuplicate = True
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

Private Sub PrintData()
Dim ii As Integer
Dim strStaffName As String '員工姓名
Dim strWeek As String '週次
    
    GetPrintLeft
    m_intPage = 0
    PrintHead
    PrintHead_1
    With Me.grd(0)
        strStaffName = .TextMatrix(1, 0)
        strWeek = .TextMatrix(1, 1)
        Printer.CurrentX = PLeft(0)
        Printer.CurrentY = m_iPrint
        Printer.Print .TextMatrix(1, 0)
        Printer.CurrentX = PLeft(1)
        Printer.CurrentY = m_iPrint
        Printer.Print .TextMatrix(1, 1)
        For ii = 1 To .Rows - 1
            '若員工姓名不同
            If strStaffName <> .TextMatrix(ii, 0) Then
                Printer.CurrentX = 0
                Printer.CurrentY = m_iPrint
                Printer.Print String(200, "-")
                m_iPrint = m_iPrint + 300
                PrintHead_1
                Printer.CurrentX = PLeft(0)
                Printer.CurrentY = m_iPrint
                Printer.Print .TextMatrix(ii, 0)
                Printer.CurrentX = PLeft(1)
                Printer.CurrentY = m_iPrint
                Printer.Print .TextMatrix(ii, 1)
                strStaffName = .TextMatrix(ii, 0)
                strWeek = .TextMatrix(ii, 1)
            '若週次不同
            ElseIf strWeek <> .TextMatrix(ii, 1) Then
                Printer.CurrentX = PLeft(1)
                Printer.CurrentY = m_iPrint
                Printer.Print String(200, "-")
                m_iPrint = m_iPrint + 300
                Printer.CurrentX = PLeft(1)
                Printer.CurrentY = m_iPrint
                Printer.Print .TextMatrix(ii, 1)
                strWeek = .TextMatrix(ii, 1)
            End If
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = m_iPrint
            Printer.Print .TextMatrix(ii, 2)
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = m_iPrint
            Printer.Print .TextMatrix(ii, 3)
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = m_iPrint
            Printer.Print .TextMatrix(ii, 4)
            Printer.CurrentX = PLeft(5)
            Printer.CurrentY = m_iPrint
            Printer.Print .TextMatrix(ii, 5)
            Printer.CurrentX = PLeft(6)
            Printer.CurrentY = m_iPrint
            Printer.Print .TextMatrix(ii, 6)
            Printer.CurrentX = PLeft(7)
            Printer.CurrentY = m_iPrint
            Printer.Print .TextMatrix(ii, 7)
            m_iPrint = m_iPrint + 300
            
            If m_iPrint > 14000 Then
                Printer.NewPage
                PrintHead
            End If
        Next ii
        Printer.CurrentX = 0
        Printer.CurrentY = m_iPrint
        Printer.Print String(200, "-")
        m_iPrint = m_iPrint + 300
        Printer.EndDoc
        ShowPrintOk
    End With
End Sub

Private Sub PrintHead()

   m_intPage = m_intPage + 1
   m_iPrint = 0: Printer.Orientation = 1
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = PLeft(3) - 250
   Printer.CurrentY = m_iPrint
   Printer.Print "工程師每週完稿明細表"
   m_iPrint = m_iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 0
   Printer.CurrentY = m_iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 4500
   Printer.CurrentY = m_iPrint
   Printer.Print "完稿年月：" & Mid(frm090625.txt1(0).Text, 1, Len(frm090625.txt1(0).Text) - 2) & "/" & Right(frm090625.txt1(0).Text, 2)
   Printer.CurrentX = 8500
   Printer.CurrentY = m_iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
   m_iPrint = m_iPrint + 300
   Printer.CurrentX = 8500
   Printer.CurrentY = m_iPrint
   Printer.Print "頁　　次：" & str(m_intPage)
   m_iPrint = m_iPrint + 300
   Printer.CurrentX = 0
   Printer.CurrentY = m_iPrint
   Printer.Print String(200, "-")
   m_iPrint = m_iPrint + 300

End Sub

Private Sub GetPrintLeft()
PLeft(0) = 0
PLeft(1) = PLeft(0) + 1125
PLeft(2) = PLeft(1) + 875
PLeft(3) = PLeft(2) + 1500
PLeft(4) = PLeft(3) + 1500
PLeft(5) = PLeft(4) + 1500
PLeft(6) = PLeft(5) + 1500
PLeft(7) = PLeft(6) + 1500
End Sub

Private Sub PrintHead_1()
    Printer.CurrentX = PLeft(0)
    Printer.CurrentY = m_iPrint
    Printer.Print "員工姓名"
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = m_iPrint
    Printer.Print "週次"
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = m_iPrint
    Printer.Print "I"
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = m_iPrint
    Printer.Print "III"
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = m_iPrint
    Printer.Print "RE"
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = m_iPrint
    Printer.Print "OT"
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = m_iPrint
    Printer.Print "PI"
    Printer.CurrentX = PLeft(7)
    Printer.CurrentY = m_iPrint
    Printer.Print "PIII"
    m_iPrint = m_iPrint + 300
   
    Printer.CurrentX = 0
    Printer.CurrentY = m_iPrint
    Printer.Print String(200, "-")
    m_iPrint = m_iPrint + 300

End Sub

Private Sub OpenWord()
Dim ii As Integer
Dim jj As Integer
Dim strSalesZone As String
Dim strSales As String
Dim intPage As Integer
Dim blnFirstPage As Boolean
Dim oTable As Word.Table

' 顯示Word程式
On Error GoTo ERRORSECTION2
    blnFirstPage = True
    If g_WordAp Is Nothing Then Set g_WordAp = New Word.Application
    g_WordAp.Documents.add
    
    g_WordAp.Visible = False
    'g_WordAp.Visible = True
    
    With g_WordAp.Application
         '加表格
         'Modified by Morgan 2018/5/23 新版 Word 可能預設是沒有框線
         '.ActiveDocument.Tables.add Range:=.Selection.Range, NumRows:=(Me.grd(0).Rows - 1) + IIf((Me.grd(0).Rows - 1) Mod 31 <> 0, Fix((Me.grd(0).Rows - 1) / 31) + 1, Fix((Me.grd(0).Rows - 1) / 31)) * 4, NumColumns:=8
         Set oTable = .ActiveDocument.Tables.add(Range:=.Selection.Range, NumRows:=(Me.grd(0).Rows - 1) + IIf((Me.grd(0).Rows - 1) Mod 31 <> 0, Fix((Me.grd(0).Rows - 1) / 31) + 1, Fix((Me.grd(0).Rows - 1) / 31)) * 4, NumColumns:=8)
         With oTable
           .Borders(wdBorderLeft).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
           .Borders(wdBorderRight).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
           .Borders(wdBorderTop).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
           .Borders(wdBorderBottom).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
           .Borders(wdBorderVertical).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
           .Borders(wdBorderHorizontal).LineStyle = g_WordAp.Options.DefaultBorderLineStyle
         End With
         'end 2018/5/23
         
        .Selection.Cells.Height = 19
        DoEvents
TitleParagraph:
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.TypeText "工程師每週完稿明細表"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCell
        '三欄
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
        .Selection.TypeText ""
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
        
        .Selection.TypeText "完稿年月：" & Mid(frm090625.txt1(0).Text, 1, Len(frm090625.txt1(0).Text) - 2) & "/" & Right(frm090625.txt1(0).Text, 2)
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        .Selection.MoveRight Unit:=wdCell
        
        .Selection.TypeText "列印日期：" & Format(strSrvDate(2), "###/##/##")
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
                
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=3, MergeBeforeSplit:=False
        .Selection.TypeText ""
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
        .Selection.MoveRight Unit:=wdCell
        intPage = intPage + 1
        .Selection.TypeText "頁　　數：" & intPage
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        .Selection.MoveRight Unit:=wdCell
            
        .Selection.SelectRow
        .Selection.Cells.Merge
        .Selection.Cells.Split NumRows:=1, NumColumns:=8, MergeBeforeSplit:=False
        .Selection.TypeText "員工姓名"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
'            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "週次"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "I"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "III"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "RE"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "OT"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "PI"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
    
        .Selection.TypeText "PIII"
        .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
        With .Selection.Cells
            .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        End With
        .Selection.MoveRight Unit:=wdCell
        
        If intPage > 1 Then GoTo ReDoFor
        strSalesZone = ""
        strSales = ""
        For ii = 1 To Me.grd(0).Rows - 1
            If blnFirstPage = False And ii Mod 31 = 1 Then
                GoTo TitleParagraph
            End If
ReDoFor:
            blnFirstPage = False
'            For jj = 0 To Me.grd(0).Cols - 1
            For jj = 0 To 7
                If jj = 0 Then
                    If strSalesZone <> Me.grd(0).TextMatrix(ii, jj) Then
                        .Selection.TypeText Me.grd(0).TextMatrix(ii, jj)
                        strSalesZone = Me.grd(0).TextMatrix(ii, jj)
                    End If
                ElseIf jj = 1 Then
                    If strSales <> Me.grd(0).TextMatrix(ii, jj) Then
                        .Selection.TypeText Me.grd(0).TextMatrix(ii, jj)
                        strSales = Me.grd(0).TextMatrix(ii, jj)
                    End If
                Else
                    .Selection.Font.Size = 9
                    .Selection.TypeText Me.grd(0).TextMatrix(ii, jj)
                    .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
                    .Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                End If
                
                If ii < Me.grd(0).Rows - 1 Then
                    .Selection.MoveRight Unit:=wdCell
                Else
'                    If jj <> Me.grd(0).Cols - 1 Then
                    If jj <> 7 Then
                        .Selection.MoveRight Unit:=wdCell
                    End If
                End If
            Next jj
        Next ii
        .Selection.WholeStory
        .Selection.Font.Name = "標楷體"
        '.Selection.WholeStory
        .Selection.Font.Name = "Courier"
    End With
    g_WordAp.Visible = True
    g_WordAp.WindowState = wdWindowStateMaximize
    MsgBox "Word檔案產生成功!!!", vbExclamation + vbOKOnly
    Exit Sub
   
ERRORSECTION2:
Select Case Err.Number
Case 91:
    'Debug.Print "ERRORSECTION2:新增一個Word 頁面"
    g_WordAp.Documents.add
    Resume Next
Case 462:
    'Debug.Print "ERRORSECTION2:新增一個Word Application物件"
    Set g_WordAp = New Word.Application
    g_WordAp.Documents.add
    Resume Next
Case Else:
    MsgBox "錯誤 : " & Err.Description, vbCritical
    Exit Sub
End Select
End Sub
