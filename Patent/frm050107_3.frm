VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm050107_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "美國IDS資料對照維護"
   ClientHeight    =   5760
   ClientLeft      =   -3375
   ClientTop       =   3030
   ClientWidth     =   9345
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9345
   Begin VB.TextBox txtChoose 
      Height          =   270
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5400
      Width           =   372
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8364
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6312
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7140
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4332
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   7646
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
      AllowUserResizing=   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
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
   Begin VB.Label Label3 
      Caption         =   "EPC,英國案發文日："
      Height          =   252
      Index           =   1
      Left            =   4320
      TabIndex        =   11
      Top             =   660
      Width           =   1692
   End
   Begin MSForms.Label lblEnginerName 
      Height          =   300
      Left            =   2160
      TabIndex        =   8
      Top             =   660
      Width           =   2172
      VariousPropertyBits=   27
      Size            =   "14499;317"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   6840
      X2              =   6960
      Y1              =   1104
      Y2              =   1104
   End
   Begin VB.Label lblDate 
      Height          =   252
      Index           =   1
      Left            =   7080
      TabIndex        =   10
      Top             =   660
      Width           =   972
   End
   Begin VB.Label lblDate 
      Height          =   252
      Index           =   0
      Left            =   6000
      TabIndex        =   9
      Top             =   660
      Width           =   972
   End
   Begin VB.Label lblEnginer 
      Height          =   252
      Left            =   1440
      TabIndex        =   7
      Top             =   660
      Width           =   852
   End
   Begin VB.Label Label2 
      Caption         =   "功能代號：           (2.修改  4.刪除  5.查詢 )"
      Height          =   252
      Left            =   120
      TabIndex        =   6
      Top             =   5400
      Width           =   3372
   End
   Begin VB.Label Label3 
      Caption         =   "美國案工程師："
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   1332
   End
End
Attribute VB_Name = "frm050107_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/3 改成Form2.0 (grdDataList,lblEnginerName)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
Private Sub cmdOK_Click(Index As Integer)
Dim intNowRow As Integer

Select Case Index
             Case 0
                     If grdDataList.Rows > 1 Then
                        intNowRow = grdDataList.row
                        frm050107_2.strCode1 = grdDataList.TextMatrix(intNowRow, 0)
                        frm050107_2.strCode2 = grdDataList.TextMatrix(intNowRow, 1)
                        frm050107_2.strCode3 = grdDataList.TextMatrix(intNowRow, 2)
                        frm050107_2.strCode4 = grdDataList.TextMatrix(intNowRow, 3)
                        frm050107_2.strCode5 = grdDataList.TextMatrix(intNowRow, 6)
                        frm050107_2.strCode6 = grdDataList.TextMatrix(intNowRow, 7)
                        frm050107_2.strCode7 = grdDataList.TextMatrix(intNowRow, 8)
                        frm050107_2.strCode8 = grdDataList.TextMatrix(intNowRow, 9)
                        frm050107_2.intChoose = Val(txtChoose)
                        frm050107_2.intWhereToGo = 1
                        frm050107_2.Show
                        Me.Hide
                     Else
                        MsgBox "資料庫無資料 !", vbInformation
                     End If
             Case 1
                        intLeaveKind = 1
                        Unload Me
             Case 2
                        intLeaveKind = 0
                        Unload Me
End Select
End Sub
Private Sub Form_Activate()
Dim varSaveCursor As Variant

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass
'edit by nickc 2007/02/05 不用 dll 了
'Set grdDataList.Recordset = obj003.ReadCaseRelationRst(lblEnginer, ChangeWDateStringToWString(lblDate(0)), ChangeWDateStringToWString(lblDate(1)), 1)
Set grdDataList.Recordset = Cls003ReadCaseRelationRst(lblEnginer, ChangeWDateStringToWString(lblDate(0)), ChangeWDateStringToWString(lblDate(1)), 1)
grdDataList.Refresh
' 90.07.16 modify by louis (加列出無檢索報告)
AppendExtData

SetDataListWidth
intLastRow = 0
If grdDataList.Rows > 1 Then
   ShowBar grdDataList, intLastRow, 12
End If
Screen.MousePointer = varSaveCursor
txtChoose.SetFocus
txtChoose = "5"
End Sub
Private Sub Form_Load()
MoveFormToCenter Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
If intLeaveKind = 1 Then
   frm050107_1.Show
Else
  Unload frm050107_1
End If
intLeaveKind = 0
'Add By Cheng 2002/07/18
Set frm050107_3 = Nothing
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 12
      blnOKtoShow = True
   End If
End If
End Sub
Private Sub grdDataList_GotFocus()
GridGotFocus grdDataList
End Sub
Private Sub grdDataList_LostFocus()
GridLostFocus grdDataList
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(400, 700, 150, 250, 2100, 750, 400, 700, 150, 250, 2100, 750, 850)
SetGridDataListWidth grdDataList, varGridWidth()
SetDataListVision grdDataList, , True
blnOKtoShow = True
End Sub
Private Sub txtChoose_GotFocus()
txtChoose.SelStart = 0
txtChoose.SelLength = Len(txtChoose)
End Sub
Private Sub txtChoose_Validate(Cancel As Boolean)
If Val(txtChoose) <> 2 And Val(txtChoose) <> 4 And Val(txtChoose) <> 5 Then
   ShowMsg MsgText(9198)
   txtChoose_GotFocus
   Cancel = True
End If
End Sub

Private Sub AppendExtData()
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim nRow As Integer
   
   strSql = "SELECT FLD1,FLD2,FLD3,FLD4,FLD5,CP14,ST02 AS FLD6,FLD7,FLD8,FLD9,FLD10,FLD11 FROM CASEPROGRESS,STAFF, " & _
               "(SELECT CM01 AS FLD1,CM02 AS FLD2,CM03 AS FLD3,CM04 AS FLD4,NVL(P1.PA05,NVL(P1.PA06,P1.PA07)) AS FLD5,C2.CP01 AS FLD7,C2.CP02 AS FLD8,C2.CP03 AS FLD9,C2.CP04 AS FLD10,NVL(P2.PA05,NVL(P2.PA06,P2.PA07)) AS FLD11,MAX(C1.CP05||C1.CP09) AS FLD12 FROM CASEPROGRESS C2,CASEMAP A,CASEPROGRESS C1,PATENT P1, PATENT P2 " & _
               "WHERE A.CM10 = '1' AND C2.CP01 = A.CM05 AND C2.CP02 = A.CM06 AND C2.CP03 = A.CM07 AND C2.CP04 = A.CM08 AND A.CM01 = C1.CP01(+) AND A.CM02 = C1.CP02(+) AND A.CM03 = C1.CP03(+) AND A.CM04 = C1.CP04(+) AND A.CM01 = P1.PA01(+) AND A.CM02 = P1.PA02(+) AND A.CM03 = P1.PA03(+) AND " & _
                     "A.CM04 = P1.PA04(+) AND C2.CP01 = P2.PA01(+) AND C2.CP02 = P2.PA02(+) AND C2.CP03 = P2.PA03(+) AND C2.CP04 = P2.PA04(+) AND " & _
               "NOT EXISTS (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS, CASEMAP B " & _
                        "WHERE B.CM10 = '1' AND B.CM05 = A.CM05 AND B.CM06 = A.CM06 AND B.CM07 = A.CM07 AND B.CM08 = A.CM08 AND CP01 = CM05 AND CP02 = CM06 AND CP03 = CM07 AND CP04 = CM08 AND CP10 = '1209') " & _
               "GROUP BY CM01,CM02,CM03,CM04,C2.CP01,C2.CP02,C2.CP03,C2.CP04,NVL(P1.PA05,NVL(P1.PA06,P1.PA07)),NVL(P2.PA05,NVL(P2.PA06,P2.PA07))) SRC1 " & _
            "WHERE CP01 = FLD1 AND CP02 = FLD2 AND CP03 = FLD3 AND CP04 = FLD4 AND CP05||CP09 = FLD12 AND CP14 = ST01(+) "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Do While rsTmp.EOF = False
         grdDataList.Rows = grdDataList.Rows + 1
         nRow = grdDataList.Rows - 1
         If IsNull(rsTmp.Fields("FLD1")) = False Then
            grdDataList.TextMatrix(nRow, 0) = rsTmp.Fields("FLD1")
         End If
         If IsNull(rsTmp.Fields("FLD2")) = False Then
            grdDataList.TextMatrix(nRow, 1) = rsTmp.Fields("FLD2")
         End If
         If IsNull(rsTmp.Fields("FLD3")) = False Then
            grdDataList.TextMatrix(nRow, 2) = rsTmp.Fields("FLD3")
         End If
         If IsNull(rsTmp.Fields("FLD4")) = False Then
            grdDataList.TextMatrix(nRow, 3) = rsTmp.Fields("FLD4")
         End If
         If IsNull(rsTmp.Fields("FLD5")) = False Then
            grdDataList.TextMatrix(nRow, 4) = rsTmp.Fields("FLD5")
         End If
         If IsNull(rsTmp.Fields("FLD6")) = False Then
            grdDataList.TextMatrix(nRow, 5) = rsTmp.Fields("FLD6")
         End If
         If IsNull(rsTmp.Fields("FLD7")) = False Then
            grdDataList.TextMatrix(nRow, 6) = rsTmp.Fields("FLD7")
         End If
         If IsNull(rsTmp.Fields("FLD8")) = False Then
            grdDataList.TextMatrix(nRow, 7) = rsTmp.Fields("FLD8")
         End If
         If IsNull(rsTmp.Fields("FLD9")) = False Then
            grdDataList.TextMatrix(nRow, 8) = rsTmp.Fields("FLD9")
         End If
         If IsNull(rsTmp.Fields("FLD10")) = False Then
            grdDataList.TextMatrix(nRow, 9) = rsTmp.Fields("FLD10")
         End If
         If IsNull(rsTmp.Fields("FLD11")) = False Then
            grdDataList.TextMatrix(nRow, 10) = rsTmp.Fields("FLD11")
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub
