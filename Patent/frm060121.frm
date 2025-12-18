VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060121 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶提供文件處理"
   ClientHeight    =   5745
   ClientLeft      =   165
   ClientTop       =   960
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9345
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Default         =   -1  'True
      Height          =   375
      Left            =   3165
      TabIndex        =   4
      Top             =   570
      Width           =   800
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm060121.frx":0000
      Left            =   1080
      List            =   "frm060121.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   5
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8388
      TabIndex        =   7
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7560
      TabIndex        =   6
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "FCP"
      Top             =   630
      Width           =   495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   1
      Top             =   630
      Width           =   855
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   2
      Top             =   630
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2640
      MaxLength       =   2
      TabIndex        =   3
      Top             =   630
      Width           =   375
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4032
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   7117
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColorBkg    =   16772048
      ScrollTrack     =   -1  'True
      FocusRect       =   2
      MergeCells      =   1
      AllowUserResizing=   1
      FormatString    =   "v|收 文 號|原文本|替換版原文本|英說|簡(繁)體中說|補文件＆資訊"
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
   Begin MSForms.Label Label8 
      Height          =   255
      Left            =   1770
      TabIndex        =   11
      Top             =   990
      Width           =   6255
      VariousPropertyBits=   27
      Size            =   "11033;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   9180
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00808080&
      X1              =   120
      X2              =   9180
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   660
      Width           =   765
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱:"
      Height          =   180
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   765
   End
End
Attribute VB_Name = "frm060121"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/10/07 改成Form2.0 ; MSHFlexGrid1改字型=新細明體-ExtB、Label8
'Create by Lydia 2018/02/01 客戶提供文件處理
Option Explicit
Dim intWhere As Integer
Dim pa(1 To 7) As String '本所案號、案件名稱
Dim intLastRow As Integer
'Added by Lydia 2019/06/27
Dim mPrevForm As Form
Dim bolOK As Boolean '直接執行確定

'Added by Lydia 2019/06/27
Public Sub SetParent(ByRef pForm As Form, ByVal pCaseNo As String)
   Set mPrevForm = pForm
   Call ChgCaseNo(pCaseNo, pa)
   If pa(2) <> "" Then
       Me.Text1 = pa(1)
       Me.Text2 = pa(2)
       Me.Text3 = pa(3)
       Me.Text4 = pa(4)
   End If
End Sub

Private Sub cmdOK_Click(Index As Integer)
   Dim i As Integer, bolChk As Boolean
   ' 記錄總收文號)
   Dim strCP09 As String
   Select Case Index
      Case 1 '確定
         bolOK = False 'Added by Lydia 2019/06/27
         With MSHFlexGrid1
            For i = 1 To .Rows - 1
               If .TextMatrix(i, 0) = "v" And "" & .TextMatrix(i, 1) <> "" Then
                  bolChk = True
                  strCP09 = .TextMatrix(i, 1)
                  Exit For
               End If
            Next
         End With
         If bolChk = False Then
            MsgBox "請選擇資料 !", vbInformation
            Exit Sub
         End If
         
         Call frm060121_1.SetParent(Me, pa(1) & pa(2) & pa(3) & pa(4), strCP09)
         frm060121_1.Show
         Me.Hide
         'Added by Lydia 2018/03/06
         If frm060121_1.ReadData = False Then
             Unload frm060121_1
             Me.Show
         End If
         'end 2018/03/06
         
      Case 2
         Unload Me
   End Select
End Sub

Private Sub Combo1_Click()
   Select Case Combo1
      Case "中"
         Label8 = pa(5)
      Case "英"
         Label8 = pa(6)
      'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
      Case "外"
         Label8 = pa(7)
   End Select
End Sub

Public Sub Command1_Click()
   If CheckCP02 = False Then Exit Sub
   Dim i As Integer
   Dim stCon As String, stNation As String
   Dim tmpBol As Boolean 'Added by Lydia 2022/06/21
   
   Label8 = ""
   Call SetGrd(True)
   
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
   
   pa(1) = Text1
   pa(2) = Text2
   pa(3) = Text3
   pa(4) = Text4

   'Added by Lydia 2022/06/21 外專後續案收文，請開放P的寰華案也可以操作
   If Text1 = "P" Then
      'Modified by Lydia 2022/07/27 開放P(非寰華案)=FMP案的收文權限
      'tmpBol = PUB_FMPtoCheck(1, 2, Pub_strUserST05, Text1, Text2, Text3, Text4)
      tmpBol = PUB_ChkIsFMP(Text1, Text2, Text3, Text4)
      If tmpBol = False Then
          'Modified by Lydia 2022/07/27 開放P(非寰華案)=FMP案的收文權限
          'MsgBox "只可收文寰華案！"
          MsgBox "只可收文寰華案／FMP案！"
          Exit Sub
      End If
   End If
   'end 2022/06/21
   
   strExc(0) = "SELECT PA05,PA06,PA07" & _
                     " FROM PATENT" & _
                     " WHERE " & ChgPatent(pa(1) & pa(2) & pa(3) & pa(4))
   intI = 0
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         For i = 0 To 2
            If IsNull(.Fields(i)) = False Then pa(i + 5) = .Fields(i)
         Next
         Label8 = pa(5)
      End With
   End If

   strExc(0) = "select ' ' v, csd05,substr(sqldatet(csd07),1,9) csd07t,substr(sqldatet(cp06),1,9) cp06t ,"
   strExc(0) = strExc(0) & "decode(csd13||csd14,null,'','Y') as F01,"
   strExc(0) = strExc(0) & "decode(csd15||csd16,null,'','Y') as F02,"
   strExc(0) = strExc(0) & "decode(csd17||csd18,null,'','Y') as F03,"
   strExc(0) = strExc(0) & "decode(csd19||csd20,null,'','Y') as F04,"
   strExc(0) = strExc(0) & "decode(csd21||csd22,null,'','優先權;')||"
   strExc(0) = strExc(0) & "decode(csd23||csd24,null,'','委任狀;')||"
   strExc(0) = strExc(0) & "decode(csd25||csd26,null,'','代表人;')||"
   strExc(0) = strExc(0) & "decode(csd27||csd28,null,'','申請人;')||"
   strExc(0) = strExc(0) & "decode(csd29||csd30,null,'','發明人;')||"
   strExc(0) = strExc(0) & "decode(csd31||csd32,null,'','非WTO;')||"
   strExc(0) = strExc(0) & "decode(csd33||csd34,null,'','其他備註') as F05_11"
   strExc(0) = strExc(0) & " from custsupportdoc,caseprogress"
   strExc(0) = strExc(0) & " where csd01='" & pa(1) & "'  and csd02='" & pa(2) & "'  and csd03='" & pa(3) & "'  and csd04='" & pa(4) & "' and nvl(csd11,0)=0 and csd05=cp09(+) "
   strExc(0) = strExc(0) & " and nvl(csd11,0)=0  order by cp06, csd05 "
   
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   
   If intI = 1 Then
       Set MSHFlexGrid1.Recordset = RsTemp
       Call SetGrd(False)
   End If
   bolOK = False 'Added by Lydia 2019/06/27
   '若只搜尋到一筆時直接勾選
   If Me.MSHFlexGrid1.Rows = 2 Then
      bolOK = True 'Added by Lydia 2019/06/27 直接執行
      MSHFlexGrid1_Click
   End If
   
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國外_FC
   Combo1.ListIndex = 0
   
   SendKeys "{Tab}"
   Call SetGrd(True)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060121 = Nothing
End Sub

Private Sub MSHFlexGrid1_Click()
   GridClick MSHFlexGrid1, intLastRow, 0
   'Modified by Lydia 2019/06/27
   'If Me.Visible = True Then cmdOK(1).SetFocus
   If Me.Visible = True Then
       If bolOK = True Then
            Call cmdOK_Click(1)
       Else
            cmdOK(1).SetFocus
       End If
   End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
    Call cmdOK_Click(1)
End Sub

Private Sub Text1_Change()
   MSHFlexGrid1.Clear
End Sub

Private Sub Text1_GotFocus()
  TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   'Modified by Lydia 2022/06/21 +P案(FMP案)
   If Text1 <> "FCP" And Text1 <> "P" Then
      MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
      TextInverse Text1
      Cancel = True
   End If
End Sub

Private Sub SetGrd(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("v", "收文號", "收件日", "本所期限", "原文本", "替換版", "英說", "簡(繁)體", "補文件＆資訊")
   arrGridHeadWidth = Array(200, 0, 800, 800, 800, 800, 800, 800, 3000)
   
   MSHFlexGrid1.Visible = False
   MSHFlexGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Rows = 2
   End If
   For iRow = 0 To MSHFlexGrid1.Cols - 1
      MSHFlexGrid1.row = 0
      MSHFlexGrid1.col = iRow
      MSHFlexGrid1.Text = arrGridHeadText(iRow)
      MSHFlexGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
   Next
   For intI = 1 To MSHFlexGrid1.Rows - 1
      MSHFlexGrid1.row = intI
      For iRow = 0 To MSHFlexGrid1.Cols - 1
         MSHFlexGrid1.col = iRow
         If iRow < 8 Then
            MSHFlexGrid1.CellAlignment = flexAlignCenterCenter
         End If
      Next iRow
   Next intI
   MSHFlexGrid1.Visible = True
End Sub

Private Sub Text2_GotFocus()
  TextInverse Text2
End Sub

'Added by Lydia 2018/03/12 跳離開,直接查詢
Private Sub Text2_LostFocus()
        If Text2 <> "" Then
           If Len(Text2) = 6 And Text2 <> pa(2) Then
                 Call Command1_Click
           ElseIf Len(Text2) <> 6 Then
                 MsgBox "本所案號請輸入6碼!! '"
                 Text2.SetFocus
                 Text2_GotFocus
           End If
        End If
End Sub

Private Sub Text3_Change()
   MSHFlexGrid1.Clear
End Sub

Private Sub Text3_GotFocus()
  TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_Change()
   MSHFlexGrid1.Clear
End Sub

Private Sub Text4_GotFocus()
  TextInverse Text4
End Sub
'檢查本所案號
Private Function CheckCP02() As Boolean
   If Len(Text2.Text) <> 6 Then
      MsgBox "本所案號輸入錯誤！"
      Text2.SetFocus
      Text2_GotFocus
      CheckCP02 = False
      Exit Function
   End If
   CheckCP02 = True
End Function

