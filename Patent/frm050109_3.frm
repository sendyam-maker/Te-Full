VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm050109_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "香港大陸案件資料維護"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9255
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7200
      TabIndex        =   2
      Top             =   60
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6360
      TabIndex        =   1
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8415
      TabIndex        =   3
      Top             =   60
      Width           =   800
   End
   Begin VB.TextBox txtChoose 
      Height          =   270
      Left            =   1110
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5325
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   4
      Top             =   165
      Width           =   492
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   3
      Left            =   2670
      MaxLength       =   2
      TabIndex        =   7
      Top             =   165
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   6
      Top             =   165
      Width           =   252
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   5
      Top             =   165
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Height          =   324
      Left            =   3180
      TabIndex        =   8
      Top             =   150
      Width           =   780
   End
   Begin VB.CommandButton Command2 
      Caption         =   "列印(&P)"
      Height          =   372
      Left            =   5505
      TabIndex        =   14
      Top             =   90
      Width           =   852
   End
   Begin VB.CommandButton Command3 
      Caption         =   "尋找(&F)"
      Height          =   324
      Left            =   3180
      TabIndex        =   13
      Top             =   495
      Width           =   780
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   5
      Left            =   1560
      MaxLength       =   6
      TabIndex        =   10
      Top             =   525
      Width           =   852
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   6
      Left            =   2400
      MaxLength       =   1
      TabIndex        =   11
      Top             =   525
      Width           =   252
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   7
      Left            =   2670
      MaxLength       =   2
      TabIndex        =   12
      Top             =   525
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   4
      Left            =   1080
      MaxLength       =   3
      TabIndex        =   9
      Top             =   525
      Width           =   492
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4065
      Left            =   150
      TabIndex        =   15
      Top             =   1185
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   7170
      _Version        =   393216
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
   Begin MSForms.Label lblEnginerName 
      Height          =   255
      Left            =   2130
      TabIndex        =   24
      Top             =   885
      Width           =   1260
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "2222;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblEnginer 
      Height          =   255
      Left            =   1440
      TabIndex        =   23
      Top             =   885
      Width           =   645
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "1138;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "大陸案發文日："
      Height          =   180
      Index           =   1
      Left            =   4710
      TabIndex        =   22
      Top             =   885
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "香港案工程師："
      Height          =   180
      Index           =   0
      Left            =   150
      TabIndex        =   21
      Top             =   885
      Width           =   1260
   End
   Begin VB.Label Label2 
      Caption         =   "功能代號：           (2.修改  4.刪除  5.查詢 )"
      Height          =   255
      Left            =   150
      TabIndex        =   20
      Top             =   5325
      Width           =   3375
   End
   Begin VB.Label lblDate 
      Height          =   255
      Index           =   0
      Left            =   6030
      TabIndex        =   19
      Top             =   885
      Width           =   975
   End
   Begin VB.Label lblDate 
      Height          =   255
      Index           =   1
      Left            =   7110
      TabIndex        =   18
      Top             =   885
      Width           =   975
   End
   Begin VB.Line Line1 
      X1              =   6870
      X2              =   6990
      Y1              =   945
      Y2              =   945
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "香港案號："
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   17
      Top             =   195
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "大陸案號："
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   16
      Top             =   540
      Width           =   900
   End
End
Attribute VB_Name = "frm050109_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; grdDataList改字型=新細明體-ExtB、lblEnginerName、lblEnginer
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
Dim iPrint As Integer, Page As Integer, PLeft(0 To 12) As Integer
Dim bolActive As Boolean '是否已啟動
'Added by Lydia 2015/07/27 +大陸澳門案(共用表單frm050109_1,frm050109_2,frm050109_3)
Public iK_CM10 As Integer  '判斷案件類別
Dim m_NA03 As String '案件-國別
Private Sub cmdOK_Click(Index As Integer)
Dim intNowRow As Integer

Select Case Index
             Case 0
                     If grdDataList.Rows > 1 Then
                        intNowRow = grdDataList.row
                        frm050109_2.strCode1 = grdDataList.TextMatrix(intNowRow, 0 + 1)
                        frm050109_2.strCode2 = grdDataList.TextMatrix(intNowRow, 1 + 1)
                        frm050109_2.strCode3 = grdDataList.TextMatrix(intNowRow, 2 + 1)
                        frm050109_2.strCode4 = grdDataList.TextMatrix(intNowRow, 3 + 1)
                        frm050109_2.strCode5 = grdDataList.TextMatrix(intNowRow, 9)
                        frm050109_2.strCode6 = grdDataList.TextMatrix(intNowRow, 10)
                        frm050109_2.strCode7 = grdDataList.TextMatrix(intNowRow, 11)
                        frm050109_2.strCode8 = grdDataList.TextMatrix(intNowRow, 12)
                        frm050109_2.intChoose = Val(txtChoose)
                        frm050109_2.iK_CM10 = iK_CM10 'Added by Lydia 2015/07/27
                        frm050109_2.intWhereToGo = 1
                        frm050109_2.Show
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

Private Sub Command1_Click()
 Dim i As Integer
   If txtCode(0) = "" Then
      MsgBox "本所案號不得空白 !", vbCritical
      txtCode(0).SetFocus
      Exit Sub
   End If
   If txtCode(1) = "" Then
      MsgBox "本所案號不得空白 !", vbCritical
      txtCode(1).SetFocus
      Exit Sub
   End If
   If txtCode(2) = "" Then txtCode(2) = "0"
   If txtCode(3) = "" Then txtCode(3) = "00"
   For i = 0 To grdDataList.Rows - 1
      If grdDataList.TextMatrix(i, 0 + 1) = txtCode(0) And grdDataList.TextMatrix(i, 1 + 1) = txtCode(1) _
         And grdDataList.TextMatrix(i, 2 + 1) = txtCode(2) And grdDataList.TextMatrix(i, 3 + 1) = txtCode(3) Then
         grdDataList.TopRow = i
         blnOKtoShow = False
         grdDataList.row = i
         ShowBar grdDataList, intLastRow, 16 + 1
         blnOKtoShow = True
         Exit For
      End If
   Next
End Sub

Private Sub Command2_Click()
On Error GoTo ErrHnd
 Dim i As Integer, j As Integer, strTxt(0 To 10) As String
   Screen.MousePointer = vbHourglass
   Page = 1
   PLeft(0) = 300
   PLeft(1) = PLeft(0) + 1600
   PLeft(2) = PLeft(1) + 2500
   PLeft(3) = PLeft(2) + 1000
   PLeft(4) = PLeft(3) + 1000
   PLeft(5) = PLeft(4) + 1600
   PLeft(6) = PLeft(5) + 2500
   PLeft(7) = PLeft(6) + 1000
   PLeft(8) = PLeft(7) + 1000
   PLeft(9) = PLeft(8) + 1000
   PLeft(10) = PLeft(9) + 1300
 
   PrintTitle
   For i = 1 To grdDataList.Rows - 1
      strTxt(0) = grdDataList.TextMatrix(i, 0 + 1) & grdDataList.TextMatrix(i, 1 + 1) & _
      grdDataList.TextMatrix(i, 2 + 1) & grdDataList.TextMatrix(i, 3 + 1)
      strTxt(1) = Left(grdDataList.TextMatrix(i, 4 + 1), 10)
      strTxt(2) = grdDataList.TextMatrix(i, 5 + 1)
      strTxt(3) = grdDataList.TextMatrix(i, 6 + 1)
      strTxt(4) = grdDataList.TextMatrix(i, 9) & grdDataList.TextMatrix(i, 10) & _
      grdDataList.TextMatrix(i, 11) & grdDataList.TextMatrix(i, 12)
      strTxt(5) = Left(grdDataList.TextMatrix(i, 13), 10)
      strTxt(6) = grdDataList.TextMatrix(i, 14)
      strTxt(7) = grdDataList.TextMatrix(i, 15)
      strTxt(8) = grdDataList.TextMatrix(i, 16)
      strTxt(9) = grdDataList.TextMatrix(i, 17)
      strTxt(10) = grdDataList.TextMatrix(i, 18)
      
      For j = 0 To 10
          Printer.CurrentX = PLeft(j)
          Printer.CurrentY = iPrint
          Printer.Print strTxt(j)
      Next
      iPrint = iPrint + 300
      
      If iPrint > 10500 Then
          Printer.NewPage
          Page = Page + 1
          PrintTitle
      End If
   Next
   Printer.EndDoc
   Screen.MousePointer = vbDefault
   MsgBox "列印結束 !", vbInformation
   Exit Sub
ErrHnd:
   Screen.MousePointer = vbDefault
   MsgBox Err.Description & " !", vbCritical
End Sub

Private Sub PrintTitle()

   iPrint = 500
   Printer.Orientation = 2
   Printer.Font.Name = "細明體"
   Printer.Font.Size = 22
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   Printer.Print "大陸香港案件資料維護表"

   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   
   If frm050109_1.txtCode(9) <> "" Then
      Printer.Print "香港案工程師：" & frm050109_1.txtCode(9)
      iPrint = iPrint + 300
   End If
   If frm050109_1.txtCode(10) <> "" Or frm050109_1.txtCode(11) <> "" Then
      Printer.Print "大陸案發文日：" & ChangeTStringToTDateString(frm050109_1.txtCode(10)) & " - " & ChangeTStringToTDateString(frm050109_1.txtCode(11))
      iPrint = iPrint + 300
   End If
   
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = 13500
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
   iPrint = iPrint + 300
   Printer.CurrentX = 13500
   Printer.CurrentY = iPrint
   Printer.Print "頁  次：" & str(Page)
   
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print "香港案號"
   
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iPrint
   Printer.Print "大陸案號"
   
   Printer.CurrentX = PLeft(5)
   Printer.CurrentY = iPrint
   Printer.Print "案件名稱"
   
   Printer.CurrentX = PLeft(6)
   Printer.CurrentY = iPrint
   Printer.Print "承辦人"
   
   Printer.CurrentX = PLeft(7)
   Printer.CurrentY = iPrint
   Printer.Print "智權人員"
   
   Printer.CurrentX = PLeft(8)
   Printer.CurrentY = iPrint
   Printer.Print "發文日"
   
   Printer.CurrentX = PLeft(9)
   Printer.CurrentY = iPrint
   Printer.Print "取消收文日"
   
   Printer.CurrentX = PLeft(10)
   Printer.CurrentY = iPrint
   Printer.Print "記錄"
   
   iPrint = iPrint + 300
   Printer.CurrentX = PLeft(0)
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
   iPrint = iPrint + 300
   
End Sub

Private Sub Command3_Click()
   Dim i As Integer
   If txtCode(4) = "" Then
      MsgBox "本所案號不得空白 !", vbCritical
      txtCode(0).SetFocus
      Exit Sub
   End If
   If txtCode(5) = "" Then
      MsgBox "本所案號不得空白 !", vbCritical
      txtCode(1).SetFocus
      Exit Sub
   End If
   If txtCode(6) = "" Then txtCode(6) = "0"
   If txtCode(7) = "" Then txtCode(7) = "00"
   For i = 1 To grdDataList.Rows - 1
      'Modified by Morgan 2018/4/26 修正欄位索引值錯誤
      If grdDataList.TextMatrix(i, 9) = txtCode(4) And grdDataList.TextMatrix(i, 10) = txtCode(5) _
         And grdDataList.TextMatrix(i, 11) = txtCode(6) And grdDataList.TextMatrix(i, 12) = txtCode(7) Then
         grdDataList.TopRow = i
         grdDataList.row = i
         blnOKtoShow = False
         ShowBar grdDataList, intLastRow, 16 + 1
         blnOKtoShow = True
         Exit For
      End If
   Next
End Sub

Private Sub Form_Activate()

If bolActive = True Then Exit Sub

Dim varSaveCursor As Variant
Dim ii As Integer

varSaveCursor = Screen.MousePointer
Screen.MousePointer = vbHourglass

'Added by Lydia 2015/07/27 +澳門大陸案
Select Case iK_CM10
    Case 4: m_NA03 = "香港"
    Case 5: m_NA03 = "澳門"
End Select
Me.Caption = "大陸" & m_NA03 & "案件資料維護"
Label3(2).Caption = m_NA03 & "案號：": Label3(0).Caption = m_NA03 & "案工程師："
'end 2015/07/27
    
'國外案已發文的也要顯示
'Modify by Morgan 2007/4/26
'Set grdDataList.Recordset = Cls003ReadCaseRelationRst(lblEnginer, ChangeTStringToWString(ChangeWDateStringToWString(lblDate(0))), ChangeTStringToWString(ChangeWDateStringToWString(lblDate(1))), 4)
'Modified by Lydia 2015/07/27
'Set grdDataList.Recordset = ReadCaseRelationRst(lblEnginer, ChangeTStringToWString(ChangeWDateStringToWString(lblDate(0))), ChangeTStringToWString(ChangeWDateStringToWString(lblDate(1))), 4)
Set grdDataList.Recordset = ReadCaseRelationRst(lblEnginer, ChangeTStringToWString(ChangeWDateStringToWString(lblDate(0))), ChangeTStringToWString(ChangeWDateStringToWString(lblDate(1))), iK_CM10)
'end 2007/4/27
grdDataList.Refresh
SetDataListWidth
intLastRow = 0
If grdDataList.Rows > 1 Then
   ShowBar grdDataList, intLastRow, 18
    For ii = 1 To Me.grdDataList.Rows - 1
      If Me.grdDataList.TextMatrix(ii, 18) = "V" Then Me.grdDataList.TextMatrix(ii, 0) = "V"
    Next ii
End If
Screen.MousePointer = varSaveCursor
txtChoose.SetFocus
txtChoose = "5"

bolActive = True
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
If intLeaveKind = 1 Then
   frm050109_1.Show
Else
  Unload frm050109_1
End If
intLeaveKind = 0
Set frm050109_3 = Nothing
End Sub

Private Sub grdDataList_DblClick()
    Dim StrSQLa As String
    With Me.grdDataList
        If Me.txtChoose.Text = "2" And Me.grdDataList.Rows > 1 Then
            If intPCaseKind = 專利 And intPWhere = 國外_CF Then
                StrSQLa = " CM01= '" & .TextMatrix(.row, 0 + 1) & "' And CM02='" & .TextMatrix(.row, 1 + 1) & "' And CM03='" & .TextMatrix(.row, 2 + 1) & "' And CM04='" & .TextMatrix(.row, 3 + 1) & "' "
                StrSQLa = StrSQLa & " And CM05= '" & .TextMatrix(.row, 9) & "' And CM06='" & .TextMatrix(.row, 10) & "' And CM07='" & .TextMatrix(.row, 11) & "' And CM08='" & .TextMatrix(.row, 12) & "' "
                'Modified by Lydia 2015/07/27
                'StrSQLa = StrSQLa & " And CM10 ='4' "
                StrSQLa = StrSQLa & " And CM10 ='" & iK_CM10 & "' "
                If .TextMatrix(.row, 18) = "" Then
                     .TextMatrix(.row, 18) = "V"
                    .TextMatrix(.row, 0) = "V"
                Else
                     .TextMatrix(.row, 18) = ""
                     
                    .TextMatrix(.row, 0) = ""
                End If
                StrSQLa = "Update CaseMap Set CM18='" & .TextMatrix(.row, 18) & "' Where " & StrSQLa
                cnnConnection.Execute StrSQLa
            End If
        End If
    End With
End Sub

Private Sub grdDataList_RowColChange()
Dim i As Integer
If intLastRow <> grdDataList.row Then
   If blnOKtoShow Then
      blnOKtoShow = False
      ShowBar grdDataList, intLastRow, 16 + 1
      For i = 0 To 3
         txtCode(i) = grdDataList.TextMatrix(grdDataList.row, i + 1)
      Next
      For i = 4 To 7
         txtCode(i) = grdDataList.TextMatrix(grdDataList.row, i + 5)
      Next
      blnOKtoShow = True
   End If
End If
End Sub
Private Sub SetDataListWidth()
Dim varGridWidth() As Variant

varGridWidth = Array(200, 400, 700, 200, 250, 1550, 750, 750, 850, 400, 700, 200, 250, 1550, 750, 750, 850, 1250, 2000)
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

Private Sub txtCode_GotFocus(Index As Integer)
   TextInverse txtCode(Index)
   Command1.Default = True
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub


