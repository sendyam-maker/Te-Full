VERSION 5.00
Begin VB.Form frm090207_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利案例資料查詢"
   ClientHeight    =   2085
   ClientLeft      =   345
   ClientTop       =   1365
   ClientWidth     =   8685
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   8685
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   0
      Left            =   900
      TabIndex        =   0
      Top             =   450
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   1
      Left            =   900
      TabIndex        =   1
      Top             =   765
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   2
      Left            =   900
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   300
      Index           =   3
      Left            =   900
      TabIndex        =   3
      Top             =   1410
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1908
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   5
      Top             =   1800
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   3888
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   7
      Top             =   1800
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   5844
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   9
      Top             =   1800
      Width           =   852
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   7776
      ScrollBars      =   3  '兩者皆有
      TabIndex        =   11
      Top             =   1800
      Width           =   852
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Index           =   0
      Left            =   1164
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1800
      Width           =   720
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Index           =   1
      Left            =   2880
      TabIndex        =   6
      Text            =   "Combo1"
      Top             =   1800
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Index           =   2
      Left            =   4836
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   1800
      Width           =   1005
   End
   Begin VB.ComboBox Combo1 
      Height          =   276
      Index           =   3
      Left            =   6852
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   1800
      Width           =   945
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   7896
      TabIndex        =   13
      Top             =   36
      Width           =   756
   End
   Begin VB.CommandButton CMDOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   7104
      TabIndex        =   12
      Top             =   36
      Width           =   756
   End
   Begin VB.Label Label8 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   2784
      TabIndex        =   22
      Top             =   1800
      Width           =   180
   End
   Begin VB.Label Label9 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   4728
      TabIndex        =   21
      Top             =   1800
      Width           =   120
   End
   Begin VB.Label Label10 
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   228
      Left            =   6720
      TabIndex        =   20
      Top             =   1800
      Width           =   132
   End
   Begin VB.Label Label11 
      Caption         =   "((("
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   852
      TabIndex        =   19
      Top             =   1800
      Width           =   300
   End
   Begin VB.Label Label9 
      Alignment       =   1  '靠右對齊
      Caption         =   "關鍵字："
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   18
      Top             =   1830
      Width           =   810
   End
   Begin VB.Label Label7 
      Alignment       =   1  '靠右對齊
      Caption         =   "備用類："
      Height          =   180
      Left            =   75
      TabIndex        =   17
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label Label5 
      Alignment       =   1  '靠右對齊
      Caption         =   "次次類："
      Height          =   180
      Left            =   75
      TabIndex        =   16
      Top             =   1140
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  '靠右對齊
      Caption         =   "次類： "
      Height          =   180
      Left            =   195
      TabIndex        =   15
      Top             =   825
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      Caption         =   "主類："
      Height          =   180
      Left            =   198
      TabIndex        =   14
      Top             =   504
      Width           =   612
   End
End
Attribute VB_Name = "frm090207_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/01/26 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit

Dim pemain As New ADODB.Recordset
Dim FINDNAME As New ADODB.Recordset
Dim SHELLSTRING As String, SSTRING As String
Dim LENSTRING As Integer, i As Integer, j As Integer
Dim a(9) As Integer, B(9) As Integer, strSQL1 As String
Public SQLSTRING As String, SQLSTRING1 As String, ADDR As Integer
'Add By Cheng 2002/03/04
Dim p As New ADODB.Recordset


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
       Case 0
         'Modify by Morgan 2004/5/10
         'strExc(0) = "SELECT (PC01||'-'||PC02||'-'||PC03||'-'||PC04) AS N1,PC09,PC11,decode(pc05,null,'',PC05||'-'||PC06||'-'||PC07||'-'||PC08) AS N2,PC10 FROM PATENTCASE  WHERE "
          strExc(0) = "SELECT (PC01||'-'||PC02||'-'||PC03||'-'||PC04) AS N1,PC09,PC11,decode(pc05,null,'',PC05||'-'||PC06||'-'||PC07||'-'||PC08) AS N2,PC10, PC20-19110000 N3, DECODE(PC19,'0','判決','1','決定書','2','其他',' ') N4 FROM PATENTCASE  WHERE PC18='1' AND "
          If Me.Combo2(0).Text = "" And Me.Combo2(1).Text = "" And Me.Combo2(2).Text = "" And Me.Combo2(3).Text = "" And txt1(0) = "" And txt1(1) = "" And txt1(2) = "" And txt1(3) = "" Then MsgBox "必須輸入查詢資料", vbInformation: Exit Sub
          'If Combo2(0).Text <> "" Then
          '  strExc(0) = strExc(0) + " PC01= '" & Combo2(0).Text & "' "
          'Else
          '  If Combo2(1).Text <> "" Then
          '      If Len(Trim(Combo2(0))) <> 0 Then
           '         strExc(0) = strExc(0) + " AND "
          '      End If
          '      strExc(0) = strExc(0) + " PC02='" & Combo2(1).Text & "' "
          '  Else
          '      If Combo2(2).Text <> "" Then
          '          If Len(Trim(Combo2(0))) <> 0 Or Len(Trim(Combo2(1))) <> 0 Then
          '              strExc(0) = strExc(0) + " AND "
          '          End If
          '          strExc(0) = strExc(0) + " PC03='" & Combo2(2).Text & "' "
          '      Else
          '         If Combo2(3).Text <> "" Then
          '              If Len(Trim(Combo2(0))) <> 0 Or Len(Trim(Combo2(1))) <> 0 Or Len(Trim(Combo2(2))) <> 0 Then
          '                  strExc(0) = strExc(0) + " AND "
          '              End If
          '              strExc(0) = strExc(0) + " PC04='" & Combo2(3).Text & "' "
          '         End If
          '      End If
          '  End If
          'End If
          If Combo2(0).Text <> "" Then
            strExc(0) = strExc(0) + " PC01= '" & Combo2(0).Text & "' "
          End If
            If Combo2(1).Text <> "" Then
                If Len(Trim(Combo2(0))) <> 0 Then
                    strExc(0) = strExc(0) + " AND "
                End If
                strExc(0) = strExc(0) + " PC02='" & Combo2(1).Text & "' "
            End If
                If Combo2(2).Text <> "" Then
                    If Len(Trim(Combo2(0))) <> 0 Or Len(Trim(Combo2(1))) <> 0 Then
                        strExc(0) = strExc(0) + " AND "
                    End If
                    strExc(0) = strExc(0) + " PC03='" & Combo2(2).Text & "' "
                End If
                   If Combo2(3).Text <> "" Then
                        If Len(Trim(Combo2(0))) <> 0 Or Len(Trim(Combo2(1))) <> 0 Or Len(Trim(Combo2(2))) <> 0 Then
                            strExc(0) = strExc(0) + " AND "
                        End If
                        strExc(0) = strExc(0) + " PC04='" & Combo2(3).Text & "' "
                   End If
          'If Combo2(1).Text <> "" Then
          '  strExc(0) = strExc(0) + "AND PC02='" & Combo2(1).Text & "' "
          'End If
          'If Combo2(2).Text <> "" Then
          '  strExc(0) = strExc(0) + "AND PC03='" & Combo2(2).Text & "' "
          'End If
          'If Combo2(3).Text <> "" Then
          '   strExc(0) = strExc(0) + "AND PC04='" & Combo2(3).Text & "' "
          'End If
          SQLSTRING1 = strExc(0)
          Me.Enabled = False
          Screen.MousePointer = vbHourglass
          Process
          Screen.MousePointer = vbDefault
          Me.Enabled = True
          'frm090207_1.Hide
          'frm090207_2.Show
       Case 1
          Unload Me
End Select
End Sub

Private Sub Combo2_Click(Index As Integer)
Select Case Index
Case 0 '主類
   If p.State <> adStateClosed Then p.Close
   strExc(1) = "SELECT DISTINCT PC02 FROM PATENTCASE Where PC01='" & Me.Combo2(0).Text & "' Order By PC02"
   p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
   If Not p.BOF Then p.MoveFirst
   Me.Combo2(Index + 1).Clear
   Do While Not p.EOF
       Combo2(Index + 1).AddItem p.Fields(0).Value
       p.MoveNext
   Loop
Case 1 '次類
   If p.State <> adStateClosed Then p.Close
   strExc(1) = "SELECT DISTINCT PC03 FROM PATENTCASE Where PC01='" & Me.Combo2(0).Text & "' And PC02='" & Me.Combo2(1).Text & "' Order By PC03"
   p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
   If Not p.BOF Then p.MoveFirst
   Me.Combo2(Index + 1).Clear
   Do While Not p.EOF
       Combo2(Index + 1).AddItem p.Fields(0).Value
       p.MoveNext
   Loop
Case 2 '次次類
   If p.State <> adStateClosed Then p.Close
   strExc(1) = "SELECT DISTINCT PC04 FROM PATENTCASE Where PC01='" & Me.Combo2(0).Text & "' And PC02='" & Me.Combo2(1).Text & "' And PC03='" & Me.Combo2(2).Text & "' Order By PC04"
   p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
   If Not p.BOF Then p.MoveFirst
   Me.Combo2(Index + 1).Clear
   Do While Not p.EOF
       Combo2(Index + 1).AddItem p.Fields(0).Value
       p.MoveNext
   Loop
End Select
End Sub

Private Sub Combo2_GotFocus(Index As Integer)
'edit by nickc 2007/07/11 切換輸入法改用API
'Me.Combo2(Index).IMEMode = 2
CloseIme
End Sub

Private Sub Combo2_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Combo2_LostFocus(Index As Integer)
Combo2_Click Index
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
'pemain.CursorLocation = adUseClient
'FINDNAME.CursorLocation = adUseClient
    Combo1(0).AddItem "", 0
    Combo1(0).AddItem "AND", 1
    Combo1(0).AddItem "NOT", 2
    Combo1(1).AddItem "", 0
    Combo1(1).AddItem "AND", 1
    Combo1(1).AddItem "OR", 2
    Combo1(1).AddItem "AND NOT", 3
    Combo1(1).AddItem "OR NOT", 4
    Combo1(2).AddItem "", 0
    Combo1(2).AddItem "AND", 1
    Combo1(2).AddItem "OR", 2
    Combo1(2).AddItem "AND NOT", 3
    Combo1(2).AddItem "OR NOT", 4
    Combo1(3).AddItem "", 0
    Combo1(3).AddItem "AND", 1
    Combo1(3).AddItem "OR", 2
    Combo1(3).AddItem "AND NOT", 3
    Combo1(3).AddItem "OR NOT", 4
    For i = 0 To 3
      Combo1(i).Text = Combo1(i).List(1)
    Next i
   'Add By Cheng 2002/03/04
   If p.State <> adStateClosed Then p.Close
   strExc(1) = "SELECT DISTINCT PC01 FROM PATENTCASE Order By PC01"
   p.Open strExc(1), cnnConnection, adOpenStatic, adLockReadOnly
   If Not p.BOF Then p.MoveFirst
   Me.Combo2(0).Clear
   Do While Not p.EOF
       Combo2(0).AddItem p.Fields(0).Value
       p.MoveNext
   Loop
End Sub

Sub Process()
'SQLSTRING = "SELECT (PC01||'-'||PC02||'-'||PC03||'-'||PC04) AS N1,PC09,PC11,(PC05||'-'||PC06||'-'||PC07||'-'||PC08) AS N2,PC10 FROM PATENTCASE WHERE "
'If Len(Trim(Text5)) <> 0 Then
'    strExc(0) = strExc(0) & " " & Replace(Replace(Text5, "A$", "PC09"), "a$", "PC09") & Replace(Replace(Text5, "A$", "PC11"), "a$", "PC11")
'End If
Screen.MousePointer = vbHourglass

ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/15 清除查詢印表記錄檔欄位
If Combo2(0).Text <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1 & Combo2(0).Text 'Add By Sindy 2010/12/15
End If
If Combo2(1).Text <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label3 & Combo2(1).Text 'Add By Sindy 2010/12/15
End If
If Combo2(2).Text <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label5 & Combo2(2).Text 'Add By Sindy 2010/12/15
End If
If Combo2(3).Text <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label7 & Combo2(3).Text 'Add By Sindy 2010/12/15
End If
If Len(Trim(txt1(0))) <> 0 Or Len(Trim(txt1(1))) <> 0 Or Len(Trim(txt1(2))) <> 0 Or Len(Trim(txt1(3))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label9(0) & Label11 & Combo1(0).Text & " " & txt1(0) & Label8 & Combo1(1).Text & " " & txt1(1) & Label9(1) & Combo1(2).Text & " " & txt1(2) & Label10 & Combo1(3).Text & " " & txt1(3) 'Add By Sindy 2010/12/15
End If

strSQL1 = "("
If Len(Trim(txt1(0))) <> 0 Then
   If Combo1(0).Text = "" Or Combo1(0).Text = "AND" Then
      strSQL1 = strSQL1 & " (instr(PC09||PC11,'" & txt1(0) & "')>0) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(0).Text & " (instr(PC09||PC11,'" & txt1(0) & "')>0 ) "
   End If
End If
If Len(Trim(txt1(1))) <> 0 Then
   If Combo1(1).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(PC09||PC11,'" & txt1(1) & "')>0 ) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(1).Text & " (instr(PC09||PC11,'" & txt1(1) & "')>0 ) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If Len(Trim(txt1(2))) <> 0 Then
   If Combo1(2).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(PC09||PC11,'" & txt1(2) & "')>0 ) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(2).Text & " (instr(PC09||PC11,'" & txt1(2) & "')>0 ) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If Len(Trim(txt1(3))) <> 0 Then
   If Combo1(3).Text = "" Then
      strSQL1 = strSQL1 & " AND (instr(PC09||PC11,'" & txt1(3) & "')>0 ) "
   Else
      strSQL1 = strSQL1 & " " & Combo1(3).Text & " (instr(PC09||PC11,'" & txt1(3) & "')>0 ) "
   End If
   strSQL1 = " (" & strSQL1 & ") "
End If
If txt1(0) <> "" Or txt1(2) <> "" Or txt1(3) <> "" Or txt1(1) <> "" Then
   If Combo2(0).Text <> "" Or Combo2(1).Text <> "" Or Combo2(2).Text <> "" Or Combo2(3).Text <> "" Then
      strSQL1 = " AND " & strSQL1
   Else
      If (InStr(1, UCase(strSQL1), "AND") > 0 And InStr(1, UCase(strSQL1), "AND") < 6) Then
         strSQL1 = Mid(strSQL1, 1, InStr(1, UCase(strSQL1), "AND") - 1) & Mid(strSQL1, InStr(1, UCase(strSQL1), "AND") + 3)
      Else
         If (InStr(1, UCase(strSQL1), "NOT") > 0 And InStr(1, UCase(strSQL1), "NOT") < 6) Then
            strSQL1 = Mid(strSQL1, 1, InStr(1, UCase(strSQL1), "NOT") - 1) & Mid(strSQL1, InStr(1, UCase(strSQL1), "NOT") + 3)
         Else
            If (InStr(1, UCase(strSQL1), "AND NOT") > 0 And InStr(1, UCase(strSQL1), "AND NOT") < 6) Then
               strSQL1 = Mid(strSQL1, 1, InStr(1, UCase(strSQL1), "AND NOT") - 1) & Mid(strSQL1, InStr(1, UCase(strSQL1), "AND NOT") + 7)
            Else
               If (InStr(1, UCase(strSQL1), "OR") > 0 And InStr(1, UCase(strSQL1), "OR") < 6) Then
                  strSQL1 = Mid(strSQL1, 1, InStr(1, UCase(strSQL1), "OR") - 1) & Mid(strSQL1, InStr(1, UCase(strSQL1), "OR") + 2)
               Else
                  If (InStr(1, UCase(strSQL1), "OR NOT") > 0 And InStr(1, UCase(strSQL1), "OR NOT") < 6) Then
                     strSQL1 = Mid(strSQL1, 1, InStr(1, UCase(strSQL1), "OR NOT") - 1) & Mid(strSQL1, InStr(1, UCase(strSQL1), "OR NOT") + 6)
                  Else
                  End If
               End If
            End If
         End If
      End If
   End If
   strExc(0) = strExc(0) & strSQL1 & ") "
Else
   'If Combo2(0).Text = "" Or Combo2(1).Text = "" Or Combo2(2).Text = "" Or Combo2(3).Text = "" Then
   '   strExc(0) = Mid(strExc(0), 1, Len(strExc(0)) - 6)
   'End If
End If

CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strExc(0), cnnConnection, adOpenStatic, adLockReadOnly
With adoRecordset
    If .RecordCount <> 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/15
        Set frm090207_2.Adodc1.Recordset = adoRecordset
        Set frm090207_2.DataGrid1.DataSource = frm090207_2.Adodc1
        frm090207_2.Command1.Enabled = True
        Me.Hide
        frm090207_2.Show
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/15
        ShowNoData
        'frm090207_2.Command1.Enabled = False
        'Screen.MousePointer = vbDefault
        Exit Sub
    End If
End With
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090207_1 = Nothing
End Sub

'Private Sub Text1_GotFocus()
'Text1.SelStart = 0
'Text1.SelLength = Len(Text1)
'Text1.IMEMode = 2
'End Sub
'
'Private Sub Text1_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'
'End Sub
'
'Private Sub Combo2(1)_GotFocus()
'Combo2(1).SelStart = 0
'Combo2(1).SelLength = Len(Combo2(1))
'Combo2(1).IMEMode = 2
'
'End Sub
'
'Private Sub Combo2(1)_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'
'End Sub
'
'Private Sub Combo2(2)_GotFocus()
'Combo2(2).SelStart = 0
'Combo2(2).SelLength = Len(Combo2(2))
'Combo2(2).IMEMode = 2
'
'End Sub
'
'Private Sub Combo2(2)_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'
'End Sub
'
'Private Sub Combo2(3)_GotFocus()
'Combo2(3).SelStart = 0
'Combo2(3).SelLength = Len(Combo2(3))
'Combo2(3).IMEMode = 2
'
'End Sub
'
'Private Sub Combo2(3)_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'
'End Sub

'Private Sub Text5_Validate(Cancel As Boolean)
'SHELLSTRING = Text5.Text
'GETSTR
'End Sub

'Private Sub GETSTR()
'A(1) = InStr(SHELLSTRING, "+")
'A(2) = InStr(SHELLSTRING, "|")
'A(3) = InStr(SHELLSTRING, "AND")
'A(4) = InStr(SHELLSTRING, "OR")
'A(5) = InStr(SHELLSTRING, "and")
'A(6) = InStr(SHELLSTRING, "or")
'Do While Not (InStr(SHELLSTRING, "+") = 0 And InStr(SHELLSTRING, "|") = 0 And InStr(SHELLSTRING, "AND") = 0 And InStr(SHELLSTRING, "OR") = 0 And InStr(SHELLSTRING, "and") = 0 And InStr(SHELLSTRING, "or") = 0)
'For i = 1 To 6
'   B(i) = A(i)
'Next i
'For i = 1 To 5
'    If B(i) > B(i + 1) Then
'       j = B(i)
'       B(i) = B(i + 1)
'       B(i + 1) = j
'    End If
'Next i
'For i = 1 To 6
'   If B(i) = 0 Then
'      j = j
'   Else
'      j = B(i): Exit For
'   End If
'Next i
'   SSTRING = Left(SHELLSTRING, j - 2)
'If j = A(1) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "AND"
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "AND"
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + ") " + "AND"
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (PC09 LIKE " + "'" + "%" + SSTRING + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + SSTRING + "%" + "'" + ") " + "AND"
'   End If
'   SHELLSTRING = Mid(SHELLSTRING, j + 2)
'
'ElseIf j = A(2) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "OR"
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "OR"
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + ") " + "OR"
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (PC09 LIKE " + "'" + "%" + SSTRING + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + SSTRING + "%" + "'" + ") " + "OR"
'   End If
'   SHELLSTRING = Mid(SHELLSTRING, j + 2)
'ElseIf j = A(3) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "AND"
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "AND"
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + ") " + "AND"
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (PC09 LIKE " + "'" + "%" + SSTRING + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + SSTRING + "%" + "'" + ") " + "AND"
'   End If
'   SHELLSTRING = Mid(SHELLSTRING, j + 4)
'ElseIf j = A(4) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "OR"
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "OR"
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + ") " + "OR"
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (PC09 LIKE " + "'" + "%" + SSTRING + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + SSTRING + "%" + "'" + ") " + "OR"
'   End If
'   SHELLSTRING = Mid(SHELLSTRING, j + 3)
'ElseIf j = A(5) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "AND"
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "AND"
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + ") " + "AND"
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (PC09 LIKE " + "'" + "%" + SSTRING + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + SSTRING + "%" + "'" + ") " + "AND"
'   End If
'   SHELLSTRING = Mid(SHELLSTRING, j + 4)

'ElseIf j = A(6) Then
'   If Mid(SSTRING, 1, 3) = "NOT" Or Mid(SSTRING, 1, 3) = "not" Or Mid(SSTRING, 1, 1) = "-" Then
'      If Mid(SSTRING, 1, 3) = "NOT" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "OR"
'      ElseIf Mid(SSTRING, 1, 3) = "not" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 5) + "%" + "'" + ") " + "OR"
'      ElseIf Mid(SSTRING, 1, 1) = "-" Then
'         SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SSTRING, 3) + "%" + "'" + ") " + "OR"
'      End If
'   Else
'      SQLSTRING = SQLSTRING + " (PC09 LIKE " + "'" + "%" + SSTRING + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + SSTRING + "%" + "'" + ") " + "OR"
''   End If
''   SHELLSTRING = Mid(SHELLSTRING, j + 3)
''End If
'A(1) = InStr(SHELLSTRING, "+")
'A(2) = InStr(SHELLSTRING, "|")
'A(3) = InStr(SHELLSTRING, "AND")
'A(4) = InStr(SHELLSTRING, "OR")
'A(5) = InStr(SHELLSTRING, "and")
'A(6) = InStr(SHELLSTRING, "or")
'Loop
'If Mid(SHELLSTRING, 1, 3) = "NOT" Or Mid(SHELLSTRING, 1, 3) = "not" Or Mid(SHELLSTRING, 1, 1) = "-" Then
'   If Mid(SHELLSTRING, 1, 3) = "NOT" Then
'      SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SHELLSTRING, 5) + "%" + "'" + " OR  " + "PC11 LIKE " + "'" + "%" + Mid(SHELLSTRING, 5) + "%" + "'" + ") "
'   ElseIf Mid(SHELLSTRING, 1, 3) = "not" Then
'      SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SHELLSTRING, 5) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SHELLSTRING, 5) + "%" + " '" + ") "
'   ElseIf Mid(SHELLSTRING, 1, 1) = "-" Then
'      SQLSTRING = SQLSTRING + "NOT" + " (PC09 LIKE " + "'" + "%" + Mid(SHELLSTRING, 3) + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + Mid(SHELLSTRING, 3) + "%" + " '" + ") "
'   End If
'Else
'   SQLSTRING = SQLSTRING + " (PC09 LIKE " + "'" + "%" + SHELLSTRING + "%" + "'" + " OR " + "PC11 LIKE " + "'" + "%" + SHELLSTRING + "%" + "'" + ") "
'End If
'End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
'edit by nickc 2007/07/11 切換輸入法改用API
'txt1(Index).IMEMode = 1
OpenIme
End Sub
