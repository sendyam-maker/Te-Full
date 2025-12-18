VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040108_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "大陸發明案件資料維護"
   ClientHeight    =   5760
   ClientLeft      =   135
   ClientTop       =   1050
   ClientWidth     =   9330
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9330
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7128
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6300
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8352
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox txtChoose 
      Height          =   270
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5400
      Width           =   372
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4272
      Left            =   120
      TabIndex        =   4
      Top             =   1020
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   7541
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      ScrollTrack     =   -1  'True
      FocusRect       =   0
      HighLight       =   0
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
      _Band(0).Cols   =   2
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   6960
      X2              =   7080
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CF 案工程師："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   1152
   End
   Begin VB.Label Label2 
      Caption         =   "功能代號：           (2.修改  4.刪除  5.查詢 )"
      Height          =   252
      Left            =   120
      TabIndex        =   10
      Top             =   5400
      Width           =   3372
   End
   Begin VB.Label lblEnginer 
      Height          =   252
      Left            =   1440
      TabIndex        =   9
      Top             =   720
      Width           =   732
   End
   Begin VB.Label lblDate 
      Height          =   252
      Index           =   0
      Left            =   5760
      TabIndex        =   8
      Top             =   720
      Width           =   972
   End
   Begin VB.Label lblDate 
      Height          =   252
      Index           =   1
      Left            =   7320
      TabIndex        =   7
      Top             =   720
      Width           =   972
   End
   Begin MSForms.Label lblEnginerName 
      Height          =   255
      Left            =   2400
      TabIndex        =   6
      Top             =   720
      Width           =   1815
      VariousPropertyBits=   27
      Size            =   "3201;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CF 案准駁日："
      Height          =   180
      Index           =   1
      Left            =   4320
      TabIndex        =   5
      Top             =   720
      Width           =   1152
   End
End
Attribute VB_Name = "frm040108_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/21 改成Form2.0 (lblEnginerName)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
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
                        frm040108_2.strCode1 = grdDataList.TextMatrix(intNowRow, 0)
                        frm040108_2.strCode2 = grdDataList.TextMatrix(intNowRow, 1)
                        frm040108_2.strCode3 = grdDataList.TextMatrix(intNowRow, 2)
                        frm040108_2.strCode4 = grdDataList.TextMatrix(intNowRow, 3)
                        frm040108_2.strCode5 = grdDataList.TextMatrix(intNowRow, 6)
                        frm040108_2.strCode6 = grdDataList.TextMatrix(intNowRow, 7)
                        frm040108_2.strCode7 = grdDataList.TextMatrix(intNowRow, 8)
                        frm040108_2.strCode8 = grdDataList.TextMatrix(intNowRow, 9)
                        frm040108_2.intChoose = Val(txtChoose)
                        frm040108_2.intWhereToGo = 1
                        frm040108_2.Show
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
'Set grdDataList.Recordset = obj003.ReadCaseRelationRst(lblEnginer, ChangeWDateStringToWString(lblDate(0)), ChangeWDateStringToWString(lblDate(1)), 2)
Set grdDataList.Recordset = Cls003ReadCaseRelationRst(lblEnginer, ChangeWDateStringToWString(lblDate(0)), ChangeWDateStringToWString(lblDate(1)), 2)
grdDataList.Refresh
SetDataListWidth
intLastRow = 0
If grdDataList.Rows > 1 Then
   ShowBar grdDataList, intLastRow, 14
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
   frm040108_1.Show
Else
  Unload frm040108_1
End If
intLeaveKind = 0
'Add By Cheng 2002/07/18
Set frm040108_3 = Nothing
End Sub
Private Sub grdDataList_RowColChange()
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 14
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

varGridWidth = Array(400, 700, 150, 250, 2100, 750, 400, 700, 150, 250, 2100, 750, 850, 1000, 1600)
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
