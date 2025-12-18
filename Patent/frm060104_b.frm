VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060104_b 
   BorderStyle     =   1  '單線固定
   Caption         =   "繳年費記錄"
   ClientHeight    =   5760
   ClientLeft      =   1740
   ClientTop       =   990
   ClientWidth     =   4890
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   4890
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   1
      Left            =   1425
      MaxLength       =   1
      TabIndex        =   7
      Top             =   945
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Index           =   0
      Left            =   1065
      MaxLength       =   7
      TabIndex        =   6
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3576
      TabIndex        =   5
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加入(&A)"
      Height          =   375
      Index           =   0
      Left            =   2355
      TabIndex        =   0
      Top             =   525
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "刪除(&D)"
      Height          =   375
      Index           =   1
      Left            =   3180
      TabIndex        =   1
      Top             =   525
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "清除(&C)"
      Height          =   375
      Index           =   2
      Left            =   4005
      TabIndex        =   2
      Top             =   525
      Width           =   800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2748
      TabIndex        =   4
      Top             =   70
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4392
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   4632
      _ExtentX        =   8176
      _ExtentY        =   7752
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
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
      _Band(0).Cols   =   3
      _Band(0).GridLinesBand=   2
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "繳費日期:"
      Height          =   180
      Index           =   87
      Left            =   225
      TabIndex        =   9
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "費用是否雙倍:         (Y:是)"
      Height          =   180
      Index           =   88
      Left            =   225
      TabIndex        =   8
      Top             =   945
      Width           =   2295
   End
End
Attribute VB_Name = "frm060104_b"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/11/17 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String, intGo As Integer, intRow As Integer, strStartDay As String, strEndDay As String
Dim pa() As String, intGo As Integer, intRow As Integer, strStartDay As String, strEndDay As String

Dim varYear As Variant '繳費年度陣列
'add by nickc 2007/02/08
Dim intWhere As Integer

Public oParent As Form 'Add by Morgan 2011/10/5


Private Sub Command2_Click(Index As Integer)
 Dim i As Integer, itmX As ListItem, strTmp As String
   If Index = 0 Then
   
      'Added by Morgan 2011/11/4 繳費年度不可大於應繳年度
      'Modified by Morgan 2012/10/11 陣列從 0 開始
      'If (MSHFlexGrid1.Rows - 1) > UBound(varYear) Then
      If (MSHFlexGrid1.Rows - 1) > (UBound(varYear) - LBound(varYear) + 1) Then
         MsgBox "繳費年度超過可繳費年度!!"
         Exit Sub
      End If
         
      strTmp = ""
      If MsgBox("是否儲存修改 ?", vbQuestion + vbYesNo) = vbYes Then
         If Not GetFeeData Then
            MsgBox "取得資料錯誤，存檔失敗 !", vbCritical
            Exit Sub
         End If
      
         strExc(1) = "UPDATE PATENT SET PA72=" & CNULL(pa(72)) & ",PA73=" & CNULL(pa(73)) & ",PA74=" & CNULL(pa(74)) & _
            " WHERE PA01='" & pa(1) & "' AND PA02='" & pa(2) & "' AND PA03='" & pa(3) & "' AND PA04='" & pa(4) & "'"
         'edit by nickc 2007/02/08 不用 dll 了
         'If objLawDll.ExecSQL(1, strExc) = True Then
         If ClsLawExecSQL(1, strExc) = True Then
            Select Case intGo
               Case 1
                  'Modify by Morgan 2011/10/5
                  'frm060104_a.Label2(9) = strTmp
                  oParent.Label2(9) = strTmp
                  
               Case 2
                  'Modify by Morgan 2007/2/9 領證繳年費退費時不設下次繳費日
                  If Me.MSHFlexGrid1.Rows > 1 Then
                     'Modified by Morgan 2012/10/11 陣列從 0 開始
                     'strTmp = Format(Val(strStartDay) + Val(varYear(MSHFlexGrid1.Rows - 1) - 1) * 10000)
                     strTmp = Format(Val(strStartDay) + Val(varYear(MSHFlexGrid1.Rows - 2)) * 10000)
                     strTmp = DBDATE(DateAdd("D", -1, ChangeWStringToWDateString(strTmp)))
                     If strTmp >= TransDate(pa(25), 2) And Val(pa(25)) > 0 Then
                        MsgBox "下次繳費日超過專用期限 (" & pa(25) & ") !", vbCritical
                        Exit Sub
                     Else
                        'Modify by Morgan 2011/10/5
                        'frm06010306_1.Text9 = TransDate(strTmp, 1)
                        oParent.Text9 = TransDate(strTmp, 1)
                     End If
                  Else
                     'Modify by Morgan 2011/10/5
                     'frm06010306_1.Text9 = ""
                     oParent.Text9 = ""
                  End If
                  'End 2007/2/9
                  
               Case 3
                  'Modify by Morgan 2006/11/14 要抓第2次繳費年-1且台灣的才要減1天
                  strTmp = Format(Val(strStartDay) + (Val(varYear(MSHFlexGrid1.Rows - 2))) * 10000)
                  If pa(9) = "000" Then
                     strTmp = CompDate(2, -1, strTmp)
                  End If
                  If strTmp >= TransDate(pa(25), 2) And Val(pa(25)) > 0 Then
                     'Modifyed by Morgan 2011/11/4 要清除有可能重複操作有殘留前次下次繳費日導致存檔仍會新增下一程序 Ex.P058380
                     'MsgBox "下次繳費日超過專用期限 (" & pa(25) & ") !", vbCritical
                     'Exit Sub
                     MsgBox "本案已無須再繳交年費 (專用期限 " & pa(25) & ") !", vbInformation, "最後繳費提醒"
                     oParent.Text9 = ""
                  Else
                     'Modify by Morgan 2011/10/5
                     'frm04010306_1.Text9 = TransDate(strTmp, 1)
                     oParent.Text9 = TransDate(strTmp, 1)
                  End If
            End Select
         Else
            MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
         End If
      End If
   End If
   Select Case intGo
      Case 1
         'Modify by Morgan 2011/10/5
         'frm060104_a.Show
         oParent.Show
      Case 2
         'Modify by Morgan 2011/10/5
         'frm06010306_1.Show
         oParent.Show
      Case 3
         'Modify by Morgan 2011/10/5
         'frm04010306_1.Show
         oParent.Show
      Case 4
         'Modify by Morgan 2011/10/5
         'frm040104_a.Show
         oParent.Show
      'Add by Morgan 2004/6/21
      Case 5
         'Modify by Morgan 2011/10/5
         'frm040104_e.Show
         oParent.Show
   End Select
   Unload Me
End Sub

Public Sub LoadMe(ByVal txt1 As String, ByVal txt2 As String, ByVal txt3 As String, ByVal txt4 As String, ByVal iGo As Integer)
 Dim strTxt(0 To 4) As String, i As Integer
   intGo = iGo
   Select Case intGo
      Case 1, 2
         intWhere = 國外_FC
      Case 3, 4
         intWhere = 國內
   End Select
   pa(1) = txt1
   pa(2) = txt2
   pa(3) = txt3
   pa(4) = txt4
   For i = 1 To 4
      strTxt(i) = pa(i)
   Next
   'edit by nickc 2007/02/08 不用 dll 了
   'If Not objPublicData.ReadPatentDatabase(pa(), intWhere, False) Then Exit Sub
   If Not ClsPDReadPatentDatabase(pa(), intWhere, False) Then Exit Sub
   
   'Add By Cheng 2002/07/17
   'Modify By Cheng 2002/10/29
   'Erase varYear
   If GetMoneyDate(Val(pa(8)), pa(9), strTxt, strStartDay, strExc(1), strEndDay) Then varYear = Split(strExc(1), ",")
   'Add by Morgan 2007/2/9 台灣未公告抓預訂公告日
   If pa(9) = "000" And strStartDay = "" Then
      strStartDay = PUB_GetPrePA14(pa)
   End If
   'End 2007/2/9
   InitFeeData
'   If objLawDll.GetNextPayDate(pa, strExc(0)) Then
'      If strExc(0) <> "" Then
'         Label2(44).Caption = Left(strExc(0), 4) - 1911 & "/" & Mid(strExc(0), 5, 2) & "/" & Mid(strExc(0), 7)
'      End If
'      Label2(47).Caption = Label2(44).Caption
'   End If
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()


   MoveFormToCenter Me
   Select Case intGo
      Case 1, 2
         intWhere = 國外_FC
      Case 3, 4
         intWhere = 國內
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060104_b = Nothing
End Sub

'將資料放入Grid
Public Sub InitFeeData()
 Dim i As Integer, varTmp1 As Variant, varTmp2 As Variant, varTmp3 As Variant, strTmp As String
 Dim strFormat As String
   GridHead
   If pa(72) <> "" Then
      varTmp1 = Split(pa(72), ",")
      varTmp2 = Split(pa(73), ",")
      varTmp3 = Split(pa(74), ",")
      With MSHFlexGrid1
         For i = 0 To UBound(varTmp1)
            strTmp = varTmp1(i)
            If UBound(varTmp2) >= i Then
               strTmp = strTmp & vbTab & TransDate(varTmp2(i), 1)
            Else
               strTmp = strTmp & vbTab & ""
            End If
            If UBound(varTmp3) >= i Then
               strTmp = strTmp & vbTab & varTmp3(i)
            Else
               strTmp = strTmp & vbTab & ""
            End If
            .AddItem strTmp, i + 1
         Next
         FixGrid MSHFlexGrid1
         If .Rows > 1 Then GridClick MSHFlexGrid1, 1, 4
      End With
   End If
End Sub

Private Sub GridHead()
   FixGrid MSHFlexGrid1
   With MSHFlexGrid1
      InitGrid 3, MSHFlexGrid1
      .Rows = 1
      .ColWidth(0) = 900:      .TextMatrix(0, 0) = "繳費年度"
      .ColWidth(1) = 900:      .TextMatrix(0, 1) = "繳費日期"
      .ColWidth(2) = 1200:     .TextMatrix(0, 2) = "費用是否雙倍"
   End With
End Sub

Private Function GetFeeData() As Boolean
 Dim i As Integer, varTemp As Variant
   For i = 72 To 74
      pa(i) = ""
   Next
   With MSHFlexGrid1
      For i = 1 To .Rows - 1
         pa(72) = pa(72) & .TextMatrix(i, 0) & ","
         pa(73) = pa(73) & TransDate(.TextMatrix(i, 1), 2) & ","
         pa(74) = pa(74) & .TextMatrix(i, 2) & ","
      Next
   End With
   For i = 72 To 74
      If Right(pa(i), 1) = "," Then pa(i) = Left(pa(i), Len(pa(i)) - 1)
   Next
   GetFeeData = True
End Function

Private Sub Command1_Click(Index As Integer)
 Dim strTmp As String, i As Integer
 Dim nPos As Integer
 Dim bFind As Boolean
On Error GoTo ErrHnd
   Select Case Index
      Case 0 '加入
'Modify by Morgan 2006/3/17
'         bFind = False
'         For nPos = 1 To MSHFlexGrid1.Rows - 1
'            If IsEmptyText(MSHFlexGrid1.TextMatrix(nPos, 0)) = True Then
'               bFind = True
'               Exit For
'            End If
'         Next nPos
'         If bFind = True Then
'            If nPos > UBound(varYear) Then
'               MsgBox "無繳費年度，無法新增資料 !", vbCritical: Exit Sub
'            Else
'               MSHFlexGrid1.TextMatrix(nPos, 0) = varYear(nPos - 1)
'            End If
'         Else
'            If MSHFlexGrid1.Rows - 1 >= UBound(varYear) Then MsgBox "無繳費年度，無法新增資料 !", vbCritical: Exit Sub
'            i = Val(varYear(MSHFlexGrid1.Rows - 1))
'            strTmp = MSHFlexGrid1.TextMatrix(MSHFlexGrid1.Rows - 1, 1)
'            'Modify By Cheng 2002/12/16
'            '不要預設繳費日期
''            If strTmp = "繳費日期" Then strTmp = strSrvDate(2)
'            If strTmp = "繳費日期" Then strTmp = ""
'            MSHFlexGrid1.AddItem Format(i) & vbTab & strTmp, MSHFlexGrid1.Rows
'            FixGrid MSHFlexGrid1
'         End If
         If MSHFlexGrid1.Rows - 2 >= UBound(varYear) Then MsgBox "無繳費年度，無法新增資料 !", vbCritical: Exit Sub
         'Add by Morgan 2010/3/10
         If pa(9) = "000" And Text3(0) = "" Then
            MsgBox "台灣案繳費日期不可空白！"
            Text3(0).SetFocus
            Exit Sub
         End If
         
         strTmp = varYear(MSHFlexGrid1.Rows - 1)
         If Text3(0) = "" Then
            strTmp = strTmp & vbTab & ""
         Else
            strTmp = strTmp & vbTab & Text3(0)
         End If
         If Text3(1) <> "" Then strTmp = strTmp & vbTab & Text3(1)
         MSHFlexGrid1.AddItem strTmp
         FixGrid MSHFlexGrid1
         MSHFlexGrid1.row = MSHFlexGrid1.Rows - 1
         If Me.MSHFlexGrid1.Rows > 1 Then GridClick MSHFlexGrid1, 1, 4
         Text3(0) = ""
         Text3(1) = ""
         Text3(0).SetFocus
'2006/3/17 end
      Case 1 '刪除
         For i = MSHFlexGrid1.row To MSHFlexGrid1.Rows - 1
            MSHFlexGrid1.RemoveItem MSHFlexGrid1.Rows - 1
         Next
         GridClick MSHFlexGrid1, MSHFlexGrid1.Rows, 4
      Case 2 '清除
         GridHead
         FixGrid MSHFlexGrid1
   End Select
   Exit Sub
ErrHnd:
   If Err.Number = 30015 Then
      GridHead
      FixGrid MSHFlexGrid1
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intRow, 4
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
   If Index = 1 Then
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub Text3_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 And Text3(0) <> "" Then Cancel = Not ChkDate(Text3(0).Text)
End Sub

