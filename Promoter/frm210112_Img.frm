VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210112_Img 
   BorderStyle     =   1  '單線固定
   Caption         =   "業務收/發文量分析_圖表"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   9315
   StartUpPosition =   3  '系統預設值
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   360
      Left            =   8070
      TabIndex        =   2
      Top             =   30
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm210112_Img.frx":0000
      Left            =   720
      List            =   "frm210112_Img.frx":000A
      Style           =   2  '單純下拉式
      TabIndex        =   1
      Top             =   195
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "print"
      Height          =   360
      Left            =   6360
      TabIndex        =   0
      Top             =   45
      Visible         =   0   'False
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5235
      Left            =   15
      TabIndex        =   5
      Top             =   555
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   9234
      _Version        =   393216
      Tabs            =   5
      Tab             =   3
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "全所"
      TabPicture(0)   =   "frm210112_Img.frx":0020
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label2"
      Tab(0).Control(1)=   "MSC1(0)"
      Tab(0).Control(2)=   "Combo2"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "區佔全所比例"
      TabPicture(1)   =   "frm210112_Img.frx":003C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "MSC1(1)"
      Tab(1).Control(2)=   "Combo3"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "區佔所有業務比例"
      TabPicture(2)   =   "frm210112_Img.frx":0058
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "MSC1(4)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "個人佔該區比例"
      TabPicture(3)   =   "frm210112_Img.frx":0074
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label4"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Combo4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "MSC1(2)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).ControlCount=   3
      TabCaption(4)   =   "個人各項比例"
      TabPicture(4)   =   "frm210112_Img.frx":0090
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "MSC1(3)"
      Tab(4).ControlCount=   1
      Begin VB.ComboBox Combo2 
         Height          =   300
         ItemData        =   "frm210112_Img.frx":00AC
         Left            =   -74235
         List            =   "frm210112_Img.frx":00BF
         Style           =   2  '單純下拉式
         TabIndex        =   7
         Top             =   390
         Width           =   1485
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   -74280
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   390
         Width           =   1425
      End
      Begin MSChart20Lib.MSChart MSC1 
         Height          =   4500
         Index           =   0
         Left            =   -74940
         OleObjectBlob   =   "frm210112_Img.frx":00D9
         TabIndex        =   8
         Top             =   690
         Width           =   9135
      End
      Begin MSChart20Lib.MSChart MSC1 
         Height          =   4455
         Index           =   1
         Left            =   -74955
         OleObjectBlob   =   "frm210112_Img.frx":242A
         TabIndex        =   9
         Top             =   690
         Width           =   9135
      End
      Begin MSChart20Lib.MSChart MSC1 
         Height          =   4830
         Index           =   3
         Left            =   -74955
         OleObjectBlob   =   "frm210112_Img.frx":477B
         TabIndex        =   10
         Top             =   360
         Width           =   9135
      End
      Begin MSChart20Lib.MSChart MSC1 
         Height          =   4440
         Index           =   2
         Left            =   60
         OleObjectBlob   =   "frm210112_Img.frx":6ACC
         TabIndex        =   11
         Top             =   705
         Width           =   9135
      End
      Begin MSChart20Lib.MSChart MSC1 
         Height          =   4830
         Index           =   4
         Left            =   -74955
         OleObjectBlob   =   "frm210112_Img.frx":8E1D
         TabIndex        =   12
         Top             =   360
         Width           =   9135
      End
      Begin MSForms.ComboBox Combo4 
         Height          =   330
         Left            =   990
         TabIndex        =   16
         Top             =   360
         Width           =   1485
         VariousPropertyBits=   679495707
         DisplayStyle    =   7
         Size            =   "2619;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "區別："
         Height          =   180
         Left            =   -74835
         TabIndex        =   15
         Top             =   435
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "所別："
         Height          =   180
         Left            =   -74865
         TabIndex        =   14
         Top             =   435
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         Height          =   180
         Left            =   90
         TabIndex        =   13
         Top             =   405
         Width           =   900
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "角度："
      Height          =   180
      Left            =   135
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "建議使用 1024 * 768 觀看"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2565
      TabIndex        =   3
      Top             =   210
      Visible         =   0   'False
      Width           =   1980
   End
End
Attribute VB_Name = "frm210112_Img"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/05 改成Form2.0 ; Combo4
'Memo By Sindy 2012/12/4 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'add by nickc 2005/08/24
Option Explicit

Dim IsBySa As Boolean
Dim Calrs As New ADODB.Recordset
Dim CuKindRs As New ADODB.Recordset
Dim DB_CuKinds As Integer
Dim DB_CuKdName() As String
'計算所別數量
Dim oTotCount As Integer
Dim oTotName As String
'計算業務區數量
Dim oSaCount As Integer
Dim oSaName As String
'計算智權人員數量
Dim oSalesCount As Integer
Dim oSalesName As String

Private Sub cmdOK_Click()
Unload Me
End Sub


Private Sub Combo1_Click()
If Combo1.Text = "智權人員" Then
   IsBySa = True
Else
   IsBySa = False
End If
'PutDataToPic
PutDataToPie
End Sub

Private Sub Combo1_Scroll()
Combo1_Click
End Sub

Private Sub Combo2_Click()
If Combo2.Text = "" Then Exit Sub
Dim arrValues()
Dim Column As Integer
Dim index1 As Integer
Dim index2 As Integer
Dim index3 As Integer
Dim index4 As Integer
Dim iNowCuCount As Integer
Dim isHaveCuKd As Boolean
Dim i As Integer
Dim tmpArr As Variant
Dim tmpCal As Integer
'全所
oTotName = ""

With MSC1(0)
   .ColumnCount = 2 'oTotCount         '所
   .RowCount = 1
   ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
   arrValues(1, 1) = "全所"
   Calrs.MoveFirst
   Do While Not Calrs.EOF
         'edit by nickc 2005/09/28
         'If CheckStr(Calrs.Fields(0).Value) = Combo2.Text Then
         If CheckStr(Calrs.Fields(0).Value) & "(" & CheckStr(Calrs.Fields(3)) & ")" = Combo2.Text Then
            If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(1)) & "'") = 0 And CheckStr(Calrs.Fields(1)) <> "" Then
               oTotName = oTotName & "'" & CheckStr(Calrs.Fields(1)) & "'"
               arrValues(1, 2) = Trim(Val(arrValues(1, 2)) + Val(Replace(Calrs.Fields("全所比例"), "%", "")))
            End If
         End If
      Calrs.MoveNext
   Loop
      arrValues(1, 3) = 100 - Val(arrValues(1, 2))
      .ChartData = arrValues
      .DataGrid.ColumnLabel(1, 1) = Combo2.Text
      .DataGrid.ColumnLabel(2, 1) = "其他所"
'      將圖表作為圖例的背景。
'先標記 , 要編譯用
      .ShowLegend = True
      .SelectPart VtChPartTypePlot, index1, index2, _
      index3, index4
      .EditCopy
      .SelectPart VtChPartTypeLegend, index1, _
      index2, index3, index4
      .EditPaste
End With


oTotName = ""
Calrs.MoveFirst
Do While Not Calrs.EOF
   'edit by nickc 2005/09/28
   'If CheckStr(Calrs.Fields(0)) = Combo2.Text Then
   If Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1) = CheckStr(Calrs.Fields(0)) Then
      If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(1)) & "(" & CheckStr(Calrs.Fields(3)) & ")" & "'") = 0 And CheckStr(Calrs.Fields(1)) <> "" Then
         oTotName = oTotName & "'" & CheckStr(Calrs.Fields(1)) & "(" & CheckStr(Calrs.Fields(3)) & ")" & "'"
      End If
   End If
   Calrs.MoveNext
Loop
Combo3.Clear
oTotName = Mid(oTotName, 2, Len(oTotName) - 2)
tmpArr = Split(oTotName, "''")
oTotCount = UBound(tmpArr) + 1
For i = 0 To UBound(tmpArr)
   Combo3.AddItem tmpArr(i), i
Next i
Combo3.ListIndex = 0


'If DB_CuKinds = 0 Then Exit Sub
'Dim arrValues()
'Dim Column As Integer
'Dim index1 As Integer
'Dim index2 As Integer
'Dim index3 As Integer
'Dim index4 As Integer
'Dim iNowCuCount As Integer
'Dim isHaveCuKd As Boolean
'Dim i As Integer
'Dim TmpArr As Variant
'oSaName = ""
'CalRs.MoveFirst
'Do While Not CalRs.EOF
'   If CheckStr(CalRs.Fields(0)) = Combo2.Text Then
'      If InStr(1, oSaName, "'" & CheckStr(CalRs.Fields(1)) & "'") = 0 And CheckStr(CalRs.Fields(1)) <> "" Then
'         oSaName = oSaName & "'" & CheckStr(CalRs.Fields(1)) & "'"
'      End If
'   End If
'   CalRs.MoveNext
'Loop
''塞預設資料
'Combo3.Clear
'oSaName = Mid(oSaName, 2, Len(oSaName) - 2)
'TmpArr = Split(oSaName, "''")
'oSaCount = UBound(TmpArr) + 1
'For i = 0 To UBound(TmpArr)
'   Combo3.AddItem TmpArr(i), i
'Next i
'Combo3.ListIndex = 0
'If IsBySa = True Then
'            '單所
'            With MSC1(1)
'               .ColumnCount = DB_CuKinds    '客戶來源種類
'               .RowCount = oSaCount          '所
'               ReDim arrValues(1 To .RowCount, 1 To DB_CuKinds)
'               For i = 0 To oSaCount - 1
'                  arrValues(i + 1, 1) = Combo3.List(i)
'               Next i
'               CalRs.MoveFirst
'               Do While Not CalRs.EOF
'                  Column = 0
'                  For i = 0 To oSaCount - 1
'                     If CheckStr(CalRs.Fields(1).Value) = Combo3.List(i) Then
'                        Column = i + 1
'                        Exit For
'                     End If
'                  Next i
'                  If Column <> 0 Then
'                     If CheckStr(CalRs.Fields(2).Value) = "區統計" Then
'                        For iNowCuCount = 2 To DB_CuKinds
'                            If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                                 Exit For
'                            End If
'                        Next iNowCuCount
'                        arrValues(Column, iNowCuCount) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'                     End If
'                  End If
'                  CalRs.MoveNext
'               Loop
'                  .ChartData = arrValues
'                  For iNowCuCount = 2 To DB_CuKinds
'                     .DataGrid.ColumnLabel(iNowCuCount - 1, 1) = DB_CuKdName(iNowCuCount)
'                  Next iNowCuCount
'
'            '      將圖表作為圖例的背景。
'            '先標記 , 要編譯用
'                  .ShowLegend = True
'                  .SelectPart VtChPartTypePlot, index1, index2, _
'                  index3, index4
'                  .EditCopy
'                  .SelectPart VtChPartTypeLegend, index1, _
'                  index2, index3, index4
'                  .EditPaste
'            End With
'Else
'            '單所
'            With MSC1(1)
'               .ColumnCount = oSaCount + 1      '所
'               .RowCount = DB_CuKinds - 1 '客戶來源種類
'               ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
'               For i = 1 To DB_CuKinds - 1
'                  arrValues(i, 1) = MidB(DB_CuKdName(i + 1), 1, 6)
'               Next i
'               CalRs.MoveFirst
'               Do While Not CalRs.EOF
'                  Column = 0
'                  For i = 0 To .ColumnCount - 1
'                     If CheckStr(CalRs.Fields(1).Value) = Combo3.List(i) Then
'                        Column = i + 1
'                        Exit For
'                     ElseIf CheckStr(CalRs.Fields(0).Value) = "" And CheckStr(CalRs.Fields(2).Value) = "全所統計" Then
'                        Column = oSaCount + 1
'                        Exit For
'                     End If
'                  Next i
'                  If Column <> 0 Then
'                     If InStr(1, CheckStr(CalRs.Fields(2).Value), "區統計") > 0 Or CheckStr(CalRs.Fields(2).Value) = "全所統計" Then
'                        For iNowCuCount = 2 To DB_CuKinds
'                            If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                                 Exit For
'                            End If
'                        Next iNowCuCount
'                        arrValues(iNowCuCount - 1, Column + 1) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'                     End If
'                  End If
'                  CalRs.MoveNext
'               Loop
'                  .ChartData = arrValues
'                  For i = 1 To oSaCount
'                     .DataGrid.ColumnLabel(i, 1) = Combo3.List(i - 1)
'                  Next i
'            .DataGrid.ColumnLabel(oSaCount + 1, 1) = "全所"
'            '      將圖表作為圖例的背景。
'            '先標記 , 要編譯用
'                  .ShowLegend = True
'                  .SelectPart VtChPartTypePlot, index1, index2, _
'                  index3, index4
'                  .EditCopy
'                  .SelectPart VtChPartTypeLegend, index1, _
'                  index2, index3, index4
'                  .EditPaste
'            End With
'End If
'Combo3_Click
End Sub

Private Sub Combo2_Scroll()
Combo2_Click
End Sub

Private Sub Combo3_Click()
If Combo3.Text = "" Then Exit Sub
Dim arrValues()
Dim Column As Integer
Dim index1 As Integer
Dim index2 As Integer
Dim index3 As Integer
Dim index4 As Integer
Dim iNowCuCount As Integer
Dim isHaveCuKd As Boolean
Dim i As Integer
Dim tmpArr As Variant
Dim tmpCal As Integer
'全所
oTotName = ""
With MSC1(1)
   .ColumnCount = 2 'oTotCount         '所
   .RowCount = 1
   ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
   arrValues(1, 1) = Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1)
   Calrs.MoveFirst
   Do While Not Calrs.EOF
         'edit by nickc 2005/09/28
         'If CheckStr(Calrs.Fields(0)) = Combo2.Text And CheckStr(Calrs.Fields(1).Value) = Combo3.Text Then
         If Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1) = CheckStr(Calrs.Fields(0)) And CheckStr(Calrs.Fields(1)) & "(" & CheckStr(Calrs.Fields(3)) & ")" = Combo3.Text Then
            If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(1)) & "'") = 0 And CheckStr(Calrs.Fields(1)) <> "" Then
               oTotName = oTotName & "'" & CheckStr(Calrs.Fields(1)) & "'"
               arrValues(1, 2) = Trim(Val(arrValues(1, 2)) + Val(Replace(Calrs.Fields("全所比例"), "%", "")))
            End If
         End If
      Calrs.MoveNext
   Loop
      arrValues(1, 3) = 100 - Val(arrValues(1, 2))
      .ChartData = arrValues
      .DataGrid.ColumnLabel(1, 1) = Combo3.Text
      .DataGrid.ColumnLabel(2, 1) = "其他區"
'      將圖表作為圖例的背景。
'先標記 , 要編譯用
      .ShowLegend = True
      .SelectPart VtChPartTypePlot, index1, index2, _
      index3, index4
      .EditCopy
      .SelectPart VtChPartTypeLegend, index1, _
      index2, index3, index4
      .EditPaste
End With

oTotName = ""
With MSC1(4)
   .ColumnCount = 2 'oTotCount         '所
   .RowCount = 1
   ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
   arrValues(1, 1) = Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1)
   Calrs.MoveFirst
   Do While Not Calrs.EOF
         'edit by nickc 2005/09/28
         'If CheckStr(Calrs.Fields(0)) = Combo2.Text And CheckStr(Calrs.Fields(1).Value) = Combo3.Text Then
         If Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1) = CheckStr(Calrs.Fields(0)) And CheckStr(Calrs.Fields(1)) & "(" & CheckStr(Calrs.Fields(3)) & ")" = Combo3.Text Then
            If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(1)) & "'") = 0 And CheckStr(Calrs.Fields(1)) <> "" Then
               oTotName = oTotName & "'" & CheckStr(Calrs.Fields(1)) & "'"
               arrValues(1, 2) = Trim(Val(arrValues(1, 2)) + Val(Replace(CheckStr(Calrs.Fields("全所業務比例")), "%", "")))
            End If
         End If
      Calrs.MoveNext
   Loop
      arrValues(1, 3) = 100 - Val(arrValues(1, 2))
      .ChartData = arrValues
      .DataGrid.ColumnLabel(1, 1) = Combo3.Text
      .DataGrid.ColumnLabel(2, 1) = "其他區"
'      將圖表作為圖例的背景。
'先標記 , 要編譯用
      .ShowLegend = True
      .SelectPart VtChPartTypePlot, index1, index2, _
      index3, index4
      .EditCopy
      .SelectPart VtChPartTypeLegend, index1, _
      index2, index3, index4
      .EditPaste
End With
oTotName = ""
Calrs.MoveFirst
Do While Not Calrs.EOF
   'edit by nickc 2005/09/28
   'If CheckStr(Calrs.Fields(0)) = Combo2.Text And CheckStr(Calrs.Fields(1)) = Combo3.Text Then
   If Mid(Combo3.Text, 1, InStr(1, Combo3.Text, "(") - 1) = CheckStr(Calrs.Fields(1)) And Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1) = CheckStr(Calrs.Fields(0)) Then
      If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(2)) & "(" & CheckStr(Calrs.Fields(3)) & ")" & "'") = 0 And CheckStr(Calrs.Fields(2)) <> "" Then
         oTotName = oTotName & "'" & CheckStr(Calrs.Fields(2)) & "(" & CheckStr(Calrs.Fields(3)) & ")" & "'"
      End If
   End If
   Calrs.MoveNext
Loop
Combo4.Clear
oTotName = Mid(oTotName, 2, Len(oTotName) - 2)
tmpArr = Split(oTotName, "''")
oTotCount = UBound(tmpArr) + 1
For i = 0 To UBound(tmpArr)
   Combo4.AddItem tmpArr(i), i
Next i
Combo4.ListIndex = 0
'Dim arrValues()
'Dim Column As Integer
'Dim index1 As Integer
'Dim index2 As Integer
'Dim index3 As Integer
'Dim index4 As Integer
'Dim iNowCuCount As Integer
'Dim isHaveCuKd As Boolean
'Dim i As Integer
'Dim TmpArr As Variant
'oSalesName = ""
'CalRs.MoveFirst
'Do While Not CalRs.EOF
'   If CheckStr(CalRs.Fields(0)) = Combo2.Text And CheckStr(CalRs.Fields(1)) = Combo3.Text Then
'      If InStr(1, oSalesName, "'" & CheckStr(CalRs.Fields(2)) & "'") = 0 And InStr(1, CheckStr(CalRs.Fields(2)), "統計") = 0 And InStr(1, CheckStr(CalRs.Fields(2)), "合計") = 0 And CheckStr(CalRs.Fields(2)) <> "" Then
'         oSalesName = oSalesName & "'" & CheckStr(CalRs.Fields(2)) & "'"
'      End If
'   End If
'   CalRs.MoveNext
'Loop
''塞預設資料
'Combo4.Clear
'oSalesName = Mid(oSalesName, 2, Len(oSalesName) - 2)
'TmpArr = Split(oSalesName, "''")
'oSalesCount = UBound(TmpArr) + 1
'For i = 0 To UBound(TmpArr)
'   Combo4.AddItem TmpArr(i), i
'Next i
'Combo4.ListIndex = 0
'If IsBySa = True Then
'
'               '單區
'               With MSC1(2)
'                  .ColumnCount = DB_CuKinds    '客戶來源種類
'                  .RowCount = oSalesCount          '區
'                  ReDim arrValues(1 To .RowCount, 1 To DB_CuKinds)
'                  For i = 0 To oSalesCount - 1
'                     arrValues(i + 1, 1) = Combo4.List(i)
'                  Next i
'                  CalRs.MoveFirst
'                  Do While Not CalRs.EOF
'                     Column = 0
'                     For i = 0 To oSalesCount - 1
'                        If CheckStr(CalRs.Fields(2).Value) = Combo4.List(i) Then
'                           Column = i + 1
'                           Exit For
'                        End If
'                     Next i
'                     If Column <> 0 Then
'                        ''If CheckStr(CalRs.Fields(2).Value) = "區統計" Then
'                           For iNowCuCount = 2 To DB_CuKinds
'                               If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                                    Exit For
'                               End If
'                           Next iNowCuCount
'                           arrValues(Column, iNowCuCount) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'                        'End If
'                     End If
'                     CalRs.MoveNext
'                  Loop
'                     .ChartData = arrValues
'                     For iNowCuCount = 2 To DB_CuKinds
'                        .DataGrid.ColumnLabel(iNowCuCount - 1, 1) = DB_CuKdName(iNowCuCount)
'                     Next iNowCuCount
'
'               '      將圖表作為圖例的背景。
'               '先標記 , 要編譯用
'                     .ShowLegend = True
'                     .SelectPart VtChPartTypePlot, index1, index2, _
'                     index3, index4
'                     .EditCopy
'                     .SelectPart VtChPartTypeLegend, index1, _
'                     index2, index3, index4
'                     .EditPaste
'               End With
'Else
'               '單區
'               With MSC1(2)
'                  .ColumnCount = oSalesCount + 1       '區
'                  .RowCount = DB_CuKinds - 1  '客戶來源種類
'                  ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
'                  For i = 1 To DB_CuKinds - 1
'                     arrValues(i, 1) = MidB(DB_CuKdName(i + 1), 1, 6)
'                  Next i
'                  CalRs.MoveFirst
'                  Do While Not CalRs.EOF
'                     Column = 0
'                     For i = 0 To .ColumnCount - 1
'                        If CheckStr(CalRs.Fields(2).Value) = Combo4.List(i) Then
'                           Column = i + 1
'                           Exit For
'                        ElseIf CheckStr(CalRs.Fields(0).Value) = "" And CheckStr(CalRs.Fields(2).Value) = "全所統計" Then
'                           Column = oSalesCount + 1
'                           Exit For
'                        End If
'                     Next i
'                     If Column <> 0 Then
'                        If CheckStr(CalRs.Fields(3).Value) <> "區統計" And CheckStr(CalRs.Fields(3).Value) <> "所有" Then
'                           For iNowCuCount = 2 To DB_CuKinds
'                               If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                                    Exit For
'                               End If
'                           Next iNowCuCount
'                           arrValues(iNowCuCount - 1, Column + 1) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'                        End If
'                     End If
'                     CalRs.MoveNext
'                  Loop
'                     .ChartData = arrValues
'                     For i = 1 To oSalesCount
'                        .DataGrid.ColumnLabel(i, 1) = Combo4.List(i - 1)
'                     Next i
'                  .DataGrid.ColumnLabel(oSalesCount + 1, 1) = "全所"
'               '      將圖表作為圖例的背景。
'               '先標記 , 要編譯用
'                     .ShowLegend = True
'                     .SelectPart VtChPartTypePlot, index1, index2, _
'                     index3, index4
'                     .EditCopy
'                     .SelectPart VtChPartTypeLegend, index1, _
'                     index2, index3, index4
'                     .EditPaste
'               End With
'End If
'Combo4_Click
End Sub

Private Sub Combo3_Scroll()
Combo3_Click
End Sub

Private Sub Combo4_Click()
If Combo4.Text = "" Then Exit Sub
Dim arrValues()
Dim Column As Integer
Dim index1 As Integer
Dim index2 As Integer
Dim index3 As Integer
Dim index4 As Integer
Dim iNowCuCount As Integer
Dim isHaveCuKd As Boolean
Dim i As Integer
Dim tmpArr As Variant
Dim tmpCal As Integer
'全所
With MSC1(2)
   .ColumnCount = 2 'oTotCount         '所
   .RowCount = 1
   ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
'   For i = 0 To Combo2.ListCount - 1
'      arrValues(i + 1, 1) = Combo2.List(i)
'   Next i
   arrValues(1, 1) = Mid(Combo3.Text, 1, InStr(1, Combo3.Text, "(") - 1)
   Calrs.MoveFirst
   Do While Not Calrs.EOF
         'edit by nickc 2005/09/28
         'If CheckStr(Calrs.Fields(0)) = Combo2.Text And CheckStr(Calrs.Fields(1)) = Combo3.Text And CheckStr(Calrs.Fields(2).Value) = Combo4.Text Then
         If Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1) = CheckStr(Calrs.Fields(0)) And Mid(Combo3.Text, 1, InStr(1, Combo3.Text, "(") - 1) = CheckStr(Calrs.Fields(1)) And CheckStr(Calrs.Fields(2).Value) & "(" & CheckStr(Calrs.Fields(3)) & ")" = Combo4.Text Then
            arrValues(1, 2) = Trim(Val(arrValues(1, 2)) + Val(Replace(Calrs.Fields("該區比例"), "%", "")))
         End If
      Calrs.MoveNext
   Loop
      arrValues(1, 3) = 100 - Val(arrValues(1, 2))
      .ChartData = arrValues
      .DataGrid.ColumnLabel(1, 1) = Combo4.Text
      .DataGrid.ColumnLabel(2, 1) = "其他同仁"
'      將圖表作為圖例的背景。
'先標記 , 要編譯用
      .ShowLegend = True
      .SelectPart VtChPartTypePlot, index1, index2, _
      index3, index4
      .EditCopy
      .SelectPart VtChPartTypeLegend, index1, _
      index2, index3, index4
      .EditPaste
End With

'算幾種種類
oTotName = ""
Calrs.MoveFirst
Do While Not Calrs.EOF
   'edit by nickc 2005/09/28
   'If CheckStr(Calrs.Fields(0)) = Combo2.Text And CheckStr(Calrs.Fields(1)) = Combo3.Text And CheckStr(Calrs.Fields(2)) = Combo4.Text Then
   If Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1) = CheckStr(Calrs.Fields(0)) And Mid(Combo3.Text, 1, InStr(1, Combo3.Text, "(") - 1) = CheckStr(Calrs.Fields(1)) And Mid(Combo4.Text, 1, InStr(1, Combo4.Text, "(") - 1) = CheckStr(Calrs.Fields(2)) Then
      If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(3)) & "'") = 0 And CheckStr(Calrs.Fields(3)) <> "" Then
         oTotName = oTotName & "'" & CheckStr(Calrs.Fields(3)) & "'"
      End If
   End If
   Calrs.MoveNext
Loop
oTotName = Mid(oTotName, 2, Len(oTotName) - 2)
tmpArr = Split(oTotName, "''")
'全所
With MSC1(3)
   .ColumnCount = UBound(tmpArr) + 1 'oTotCount         '所
   .RowCount = 1
   ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
   arrValues(1, 1) = Mid(Combo4.Text, 1, InStr(1, Combo4.Text, "(") - 1)
   Calrs.MoveFirst
   Do While Not Calrs.EOF
      'edit by nickc 2005/09/28
      'If CheckStr(Calrs.Fields(0)) = Combo2.Text And CheckStr(Calrs.Fields(1)) = Combo3.Text And CheckStr(Calrs.Fields(2)) = Combo4.Text Then
      'If CheckStr(Calrs.Fields(0)) & "(" & CheckStr(Calrs.Fields(3)) & ")" = Combo2.Text And CheckStr(Calrs.Fields(1)) = Combo3.Text And CheckStr(Calrs.Fields(2)) = Combo4.Text Then
      If Mid(Combo2.Text, 1, InStr(1, Combo2.Text, "(") - 1) = CheckStr(Calrs.Fields(0)) And Mid(Combo3.Text, 1, InStr(1, Combo3.Text, "(") - 1) = CheckStr(Calrs.Fields(1)) And Mid(Combo4.Text, 1, InStr(1, Combo4.Text, "(") - 1) = CheckStr(Calrs.Fields(2)) Then
         For i = 0 To UBound(tmpArr)
            If CheckStr(Calrs.Fields(3).Value) = tmpArr(i) Then
               arrValues(1, 2 + i) = Trim(Val(arrValues(1, 2 + i)) + Val(Replace(Calrs.Fields("個人比例"), "%", "")))
               Exit For
            End If
         Next i
      End If
      Calrs.MoveNext
   Loop
      
      .ChartData = arrValues
      For i = 0 To UBound(tmpArr)
         .DataGrid.ColumnLabel(i + 1, 1) = tmpArr(i)
      Next i
'      將圖表作為圖例的背景。
'先標記 , 要編譯用
      .ShowLegend = True
      .SelectPart VtChPartTypePlot, index1, index2, _
      index3, index4
      .EditCopy
      .SelectPart VtChPartTypeLegend, index1, _
      index2, index3, index4
      .EditPaste
End With










'Dim arrValues()
'Dim Column As Integer
'Dim index1 As Integer
'Dim index2 As Integer
'Dim index3 As Integer
'Dim index4 As Integer
'Dim iNowCuCount As Integer
'Dim isHaveCuKd As Boolean
'Dim i As Integer
'If IsBySa = True Then
'            '單業務
'            With MSC1(3)
'               .ColumnCount = DB_CuKinds    '客戶來源種類
'               .RowCount = 1          '業務
'               ReDim arrValues(1 To .RowCount, 1 To DB_CuKinds)
'               For i = 0 To 1 - 1
'                  arrValues(i + 1, 1) = Combo4.Text
'               Next i
'               CalRs.MoveFirst
'               Do While Not CalRs.EOF
'                  Column = 0
'                  For i = 0 To 1 - 1
'                     If CheckStr(CalRs.Fields(2).Value) = Combo4.Text Then
'                        Column = i + 1
'                        Exit For
'                     End If
'                  Next i
'                  If Column <> 0 Then
'                     'If CheckStr(CalRs.Fields(3).Value) = "區統計" Then
'                        For iNowCuCount = 2 To DB_CuKinds
'                            If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                                 Exit For
'                            End If
'                        Next iNowCuCount
'                        arrValues(Column, iNowCuCount) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'                     'End If
'                  End If
'                  CalRs.MoveNext
'               Loop
'                  .ChartData = arrValues
'                  For iNowCuCount = 2 To DB_CuKinds
'                     .DataGrid.ColumnLabel(iNowCuCount - 1, 1) = DB_CuKdName(iNowCuCount)
'                  Next iNowCuCount
'
'            '      將圖表作為圖例的背景。
'            '先標記 , 要編譯用
'                  .ShowLegend = True
'                  .SelectPart VtChPartTypePlot, index1, index2, _
'                  index3, index4
'                  .EditCopy
'                  .SelectPart VtChPartTypeLegend, index1, _
'                  index2, index3, index4
'                  .EditPaste
'            End With
'Else
'            '單業務
'            With MSC1(3)
'               .ColumnCount = 2            '業務
'               .RowCount = DB_CuKinds - 1 '客戶來源種類
'               ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
'               For i = 1 To DB_CuKinds - 1
'                  arrValues(i, 1) = MidB(DB_CuKdName(i + 1), 1, 6)
'               Next i
'               CalRs.MoveFirst
'               Do While Not CalRs.EOF
'                  Column = 0
'                  For i = 0 To .ColumnCount - 1
'                     If CheckStr(CalRs.Fields(2).Value) = Combo4.Text Then
'                        Column = i + 1
'                        Exit For
'                     ElseIf CheckStr(CalRs.Fields(0).Value) = "" And CheckStr(CalRs.Fields(2).Value) = "全所統計" Then
'                        Column = 2
'                     End If
'                  Next i
'                  If Column <> 0 Then
'                     If CheckStr(CalRs.Fields(3).Value) <> "區統計" And CheckStr(CalRs.Fields(3).Value) <> "所有" Then
'                        For iNowCuCount = 2 To DB_CuKinds
'                            If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                                 Exit For
'                            End If
'                        Next iNowCuCount
'                        arrValues(iNowCuCount - 1, Column + 1) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'                     End If
'                  End If
'                  CalRs.MoveNext
'               Loop
'                  .ChartData = arrValues
'                  .DataGrid.ColumnLabel(1, 1) = Combo4.Text
'                  .DataGrid.ColumnLabel(2, 1) = "全所"
'            '      將圖表作為圖例的背景。
'            '先標記 , 要編譯用
'                  .ShowLegend = True
'                  .SelectPart VtChPartTypePlot, index1, index2, _
'                  index3, index4
'                  .EditCopy
'                  .SelectPart VtChPartTypeLegend, index1, _
'                  index2, index3, index4
'                  .EditPaste
'            End With
'
'
'End If
End Sub

Private Sub Combo4_Scroll()
Combo4_Click
End Sub

Private Sub Command1_Click()
'測試印圖
Pub_Can_Copy_Pic = True
MSC1(SSTab1.Tab).EditCopy
Printer.Orientation = 2
Printer.Print " "
Printer.PaintPicture Clipboard.GetData(), 0, 0, Printer.ScaleWidth, Printer.ScaleHeight
Printer.EndDoc
Pub_Can_Copy_Pic = False
End Sub

Private Sub Form_Load()

   
'frm210112_Img.Width = mdiMain.ScaleWidth
'frm210112_Img.Left = 50
'frm210112_Img.Height = mdiMain.ScaleHeight
'frm210112_Img.Top = Screen.Height - mdiMain.Height - 50
'frm210112_Img.SSTab1.Width = frm210112_Img.ScaleWidth - 50
'frm210112_Img.SSTab1.Height = frm210112_Img.ScaleHeight - 600
'frm210112_Img.MSC1(0).Height = frm210112_Img.SSTab1.Height - 400
'frm210112_Img.MSC1(0).Width = frm210112_Img.SSTab1.Width - 400
'frm210112_Img.MSC1(1).Height = frm210112_Img.SSTab1.Height - 400
'frm210112_Img.MSC1(1).Width = frm210112_Img.SSTab1.Width - 400
'frm210112_Img.MSC1(2).Height = frm210112_Img.SSTab1.Height - 400
'frm210112_Img.MSC1(2).Width = frm210112_Img.SSTab1.Width - 400
'frm210112_Img.MSC1(3).Height = frm210112_Img.SSTab1.Height - 400
'frm210112_Img.MSC1(3).Width = frm210112_Img.SSTab1.Width - 400
'frm210112_Img.SSTab1.Tab = 0
'frm210112_Img.Label2.Left = 500
'frm210112_Img.Combo2.Left = frm210112_Img.Label3.Left + 700
'frm210112_Img.SSTab1.Tab = 1
'frm210112_Img.Label3.Left = (frm210112_Img.SSTab1.Width / 4) + 500
'frm210112_Img.Combo3.Left = frm210112_Img.Label3.Left + 700
'frm210112_Img.SSTab1.Tab = 2
'frm210112_Img.Label4.Left = (frm210112_Img.SSTab1.Width / 2) + 500
'frm210112_Img.Combo4.Left = frm210112_Img.Label4.Left + 700
'frm210112_Img.SSTab1.Tab = 3

Label2.Visible = False
Combo2.Visible = False
Label3.Visible = False
Combo3.Visible = False
Label4.Visible = False
Combo4.Visible = False
If frm210112.txtSalesArea1.Enabled = False And frm210112.txtSalesArea.Enabled = False Then
   SSTab1.TabEnabled(0) = False
   SSTab1.Tab = 1
End If
If frm210112.txtSales.Enabled = False Then
   SSTab1.TabEnabled(1) = False
   SSTab1.Tab = 2
   Combo4.Enabled = False
End If
Select Case SSTab1.Tab
Case 0
      Label2.Visible = True
      Combo2.Visible = True
      Label3.Visible = False
      Combo3.Visible = False
      Label4.Visible = False
      Combo4.Visible = False
Case 1, 2
      Label2.Visible = False
      Combo2.Visible = False
      Label3.Visible = True
      Combo3.Visible = True
      Label4.Visible = False
      Combo4.Visible = False
Case 3
      Label2.Visible = False
      Combo2.Visible = False
      Label3.Visible = False
      Combo3.Visible = False
      Label4.Visible = True
      Combo4.Visible = True
Case 4
      Label2.Visible = False
      Combo2.Visible = False
      Label3.Visible = False
      Combo3.Visible = False
      Label4.Visible = False
      Combo4.Visible = False
Case Else
End Select
MoveFormToCenter Me
Combo1.Text = Combo1.List(0)
Combo2.Clear
Combo3.Clear
Combo4.Clear
'add by nickc 2005/09/21
PutDataToPie
'PutDataToPic
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm210112_Img = Nothing
End Sub

''依智權人員
'Sub PutDataToPicBySa()
'Dim arrValues()
'Dim Column As Integer
'Dim index1 As Integer
'Dim index2 As Integer
'Dim index3 As Integer
'Dim index4 As Integer
'Dim iNowCuCount As Integer
'Dim isHaveCuKd As Boolean
'Dim i As Integer
''全所
'With MSC1(0)
'   .ColumnCount = DB_CuKinds    '客戶來源種類
'   .RowCount = oTotCount + 1          '所
'   ReDim arrValues(1 To .RowCount, 1 To DB_CuKinds)
'   For i = 0 To oTotCount - 1
'      arrValues(i + 1, 1) = Combo2.List(i)
'   Next i
'   arrValues(.RowCount, 1) = "全所"
'   CalRs.MoveFirst
'   Do While Not CalRs.EOF
'      Column = 0
'      For i = 0 To oTotCount - 1
'         If CheckStr(CalRs.Fields(0).Value) = Combo2.List(i) Then
'            Column = i + 1
'            Exit For
'         ElseIf CheckStr(CalRs.Fields(0).Value) = "" And CheckStr(CalRs.Fields(2).Value) = "全所統計" Then
'            Column = oTotCount + 1
'            Exit For
'         End If
'      Next i
'      If Column <> 0 Then
'         If CheckStr(CalRs.Fields(1).Value) = "" And CheckStr(CalRs.Fields(3).Value) <> "所有" Then
'            For iNowCuCount = 2 To DB_CuKinds
'                If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                     Exit For
'                End If
'            Next iNowCuCount
'            arrValues(Column, iNowCuCount) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'         End If
'      End If
'      CalRs.MoveNext
'   Loop
'      .ChartData = arrValues
'      For iNowCuCount = 2 To DB_CuKinds
'         .DataGrid.ColumnLabel(iNowCuCount - 1, 1) = DB_CuKdName(iNowCuCount)
'      Next iNowCuCount
'      For i = 1 To oTotCount
'         .DataGrid.RowLabel(i, 1) = Combo2.List(i - 1)
'      Next i
'      .DataGrid.RowLabel(oTotCount + 1, 1) = "全所"
''      將圖表作為圖例的背景。
''先標記 , 要編譯用
'      .ShowLegend = True
'      .SelectPart VtChPartTypePlot, index1, index2, _
'      index3, index4
'      .EditCopy
'      .SelectPart VtChPartTypeLegend, index1, _
'      index2, index3, index4
'      .EditPaste
'End With
'Combo2_Click
'
'End Sub

'依案件來源
'Sub PutDataToPicByCS()
'Dim arrValues()
'Dim Column As Integer
'Dim index1 As Integer
'Dim index2 As Integer
'Dim index3 As Integer
'Dim index4 As Integer
'Dim iNowCuCount As Integer
'Dim isHaveCuKd As Boolean
'Dim i As Integer
''全所
'With MSC1(0)
'   .ColumnCount = oTotCount + 1        '所
'   .RowCount = DB_CuKinds - 1 '客戶來源種類
'   ReDim arrValues(1 To .RowCount, 1 To .ColumnCount + 1)
'   For i = 1 To DB_CuKinds - 1
'      arrValues(i, 1) = MidB(DB_CuKdName(i + 1), 1, 6)
'   Next i
'   CalRs.MoveFirst
'   Do While Not CalRs.EOF
'      Column = 0
'      For i = 0 To .ColumnCount - 1
'         If CheckStr(CalRs.Fields(0).Value) = Combo2.List(i) Then
'            Column = i + 1
'            Exit For
'         ElseIf CheckStr(CalRs.Fields(0).Value) = "" And CheckStr(CalRs.Fields(2).Value) = "全所統計" Then
'            Column = oTotCount + 1
'            Exit For
'         End If
'      Next i
'      If Column <> 0 Then
'         If CheckStr(CalRs.Fields(1).Value) = "" And CheckStr(CalRs.Fields(3).Value) <> "所有" Then
'            For iNowCuCount = 2 To DB_CuKinds
'                If CheckStr(CalRs.Fields(3).Value) = DB_CuKdName(iNowCuCount) Then
'                     Exit For
'                End If
'            Next iNowCuCount
'            arrValues(iNowCuCount - 1, Column + 1) = CalRs.Fields(4).Value ' Series 1 values.    '第一種客戶來源
'         End If
'      End If
'      CalRs.MoveNext
'   Loop
'      .ChartData = arrValues
'      For i = 1 To oTotCount
'         .DataGrid.ColumnLabel(i, 1) = Combo2.List(i - 1)
'      Next i
'      .DataGrid.ColumnLabel(oTotCount + 1, 1) = "全所"
'      For i = 1 To DB_CuKinds - 1
'         .DataGrid.RowLabel(i, 1) = MidB(DB_CuKdName(i + 1), 1, 6)
'      Next i
''      將圖表作為圖例的背景。
''先標記 , 要編譯用
'      .ShowLegend = True
'      .SelectPart VtChPartTypePlot, index1, index2, _
'      index3, index4
'      .EditCopy
'      .SelectPart VtChPartTypeLegend, index1, _
'      index2, index3, index4
'      .EditPaste
'End With
'Combo2_Click
'End Sub

'Sub PutDataToPic()
'Dim i As Integer
'Screen.MousePointer = vbHourglass
'If frm210112.opt1(0).Value = True Then
'   MSC1(0).ChartType = VtChChartType2dBar
'   MSC1(1).ChartType = VtChChartType2dBar
'   MSC1(2).ChartType = VtChChartType2dBar
'   MSC1(3).ChartType = VtChChartType2dBar
'Else
'   MSC1(0).ChartType = VtChChartType2dPie
'   MSC1(1).ChartType = VtChChartType2dPie
'   MSC1(2).ChartType = VtChChartType2dPie
'   MSC1(3).ChartType = VtChChartType2dPie
'End If
'Set CalRs = frm210112.grdDataList.Recordset.Clone
'
'Dim TmpArr As Variant
'oTotName = ""
'CalRs.MoveFirst
'Do While Not CalRs.EOF
'   If InStr(1, oTotName, "'" & CheckStr(CalRs.Fields(0)) & "'") = 0 And CheckStr(CalRs.Fields(0)) <> "" Then
'      oTotName = oTotName & "'" & CheckStr(CalRs.Fields(0)) & "'"
'   End If
'   CalRs.MoveNext
'Loop
''塞預設資料
'Combo2.Clear
'oTotName = Mid(oTotName, 2, Len(oTotName) - 2)
'TmpArr = Split(oTotName, "''")
'oTotCount = UBound(TmpArr) + 1
'For i = 0 To UBound(TmpArr)
'   Combo2.AddItem TmpArr(i), i
'Next i
'Combo2.ListIndex = 0
'strSQL = "select * from casesourcemap order by csm01 "
'Set CuKindRs = New ADODB.Recordset
'If CuKindRs.State = 1 Then CuKindRs.Close
'CuKindRs.CursorLocation = adUseClient
'CuKindRs.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'If CuKindRs.RecordCount <> 0 Then
'   DB_CuKinds = CuKindRs.RecordCount + 1
'   ReDim DB_CuKdName(2 To DB_CuKinds) As String
'   CuKindRs.MoveFirst
'   Do While Not CuKindRs.EOF
'      DB_CuKdName(CuKindRs.AbsolutePosition + 1) = CheckStr(CuKindRs.Fields(1).Value)
'      CuKindRs.MoveNext
'   Loop
'Else
'   DB_CuKinds = 1
'End If
'Set CuKindRs = Nothing
'If IsBySa = True Then
'   PutDataToPicBySa
'Else
'   PutDataToPicByCS
'End If
'Screen.MousePointer = vbDefault
'End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Select Case SSTab1.Tab
Case 0
      Label2.Visible = True
      Combo2.Visible = True
      Label3.Visible = False
      Combo3.Visible = False
      Label4.Visible = False
      Combo4.Visible = False
Case 1, 2
      Label2.Visible = False
      Combo2.Visible = False
      Label3.Visible = True
      Combo3.Visible = True
      Label4.Visible = False
      Combo4.Visible = False
Case 3
      Label2.Visible = False
      Combo2.Visible = False
      Label3.Visible = False
      Combo3.Visible = False
      Label4.Visible = True
      Combo4.Visible = True
Case 4
      Label2.Visible = False
      Combo2.Visible = False
      Label3.Visible = False
      Combo3.Visible = False
      Label4.Visible = False
      Combo4.Visible = False
Case Else
End Select
End Sub
'add by nickc 2005/09/21 畫圓餅圖
Function PutDataToPie()
Screen.MousePointer = vbHourglass
   MSC1(0).ChartType = VtChChartType2dPie
   MSC1(1).ChartType = VtChChartType2dPie
   MSC1(2).ChartType = VtChChartType2dPie
   MSC1(3).ChartType = VtChChartType2dPie
   MSC1(4).ChartType = VtChChartType2dPie
Dim arrValues()
Dim Column As Integer
Dim index1 As Integer
Dim index2 As Integer
Dim index3 As Integer
Dim index4 As Integer
Dim iNowCuCount As Integer
Dim isHaveCuKd As Boolean
Dim i As Integer
'計算幾個所
Set Calrs = frm210112.Calrs.Clone

Dim tmpArr As Variant
oTotName = ""
Calrs.MoveFirst
Do While Not Calrs.EOF
'edit by nickc 2005/09/28
'   If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(0)) & "'") = 0 And CheckStr(Calrs.Fields(0)) <> "" Then
'      oTotName = oTotName & "'" & CheckStr(Calrs.Fields(0)) & "'"
   If InStr(1, oTotName, "'" & CheckStr(Calrs.Fields(0)) & "(" & CheckStr(Calrs.Fields(3)) & ")" & "'") = 0 And CheckStr(Calrs.Fields(0)) <> "" Then
      oTotName = oTotName & "'" & CheckStr(Calrs.Fields(0)) & "(" & CheckStr(Calrs.Fields(3)) & ")" & "'"
      'Combo2.AddItem
   End If
   Calrs.MoveNext
Loop
Combo2.Clear
oTotName = Mid(oTotName, 2, Len(oTotName) - 2)
tmpArr = Split(oTotName, "''")
oTotCount = UBound(tmpArr) + 1
For i = 0 To UBound(tmpArr)
   Combo2.AddItem tmpArr(i), i
Next i
Combo2.ListIndex = 0

Screen.MousePointer = vbDefault
End Function




