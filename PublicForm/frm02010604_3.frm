VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm02010604_3 
   BorderStyle     =   1  '單線固定
   Caption         =   "分割案件關係維護"
   ClientHeight    =   5745
   ClientLeft      =   105
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
   Begin VB.TextBox txtCode 
      Enabled         =   0   'False
      Height          =   264
      Index           =   4
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   6
      Top             =   210
      Visible         =   0   'False
      Width           =   252
   End
   Begin VB.CommandButton Command2 
      Caption         =   "列印(&P)"
      Height          =   372
      Left            =   5472
      TabIndex        =   18
      Top             =   96
      Visible         =   0   'False
      Width           =   852
   End
   Begin VB.CommandButton Command1 
      Caption         =   "尋找(&F)"
      Height          =   324
      Left            =   3270
      TabIndex        =   9
      Top             =   192
      Width           =   780
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   1
      Left            =   1536
      MaxLength       =   6
      TabIndex        =   5
      Top             =   216
      Width           =   1005
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   2
      Left            =   2550
      MaxLength       =   1
      TabIndex        =   7
      Top             =   216
      Width           =   252
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   3
      Left            =   2820
      MaxLength       =   2
      TabIndex        =   8
      Top             =   216
      Width           =   372
   End
   Begin VB.TextBox txtCode 
      Height          =   264
      Index           =   0
      Left            =   1056
      MaxLength       =   3
      TabIndex        =   4
      Top             =   216
      Width           =   492
   End
   Begin VB.TextBox txtChoose 
      Height          =   270
      Left            =   1080
      MaxLength       =   1
      TabIndex        =   0
      Top             =   5340
      Width           =   372
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   2
      Left            =   8388
      TabIndex        =   3
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   6336
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   7164
      TabIndex        =   2
      Top             =   70
      Width           =   1200
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4152
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   9072
      _ExtentX        =   16007
      _ExtentY        =   7329
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "分割案號："
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   240
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   7200
      X2              =   7320
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Label lblDate 
      Height          =   255
      Index           =   1
      Left            =   7410
      TabIndex        =   16
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label lblDate 
      Height          =   255
      Index           =   0
      Left            =   5880
      TabIndex        =   15
      Top             =   720
      Width           =   1245
   End
   Begin VB.Label lblEnginer 
      Height          =   255
      Left            =   1050
      TabIndex        =   14
      Top             =   720
      Width           =   2865
   End
   Begin VB.Label Label2 
      Caption         =   "功能代號：           (2.修改  4.刪除  5.查詢 )"
      Height          =   252
      Left            =   120
      TabIndex        =   13
      Top             =   5340
      Width           =   3372
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請人編號："
      Height          =   180
      Index           =   1
      Left            =   4680
      TabIndex        =   10
      Top             =   720
      Width           =   1080
   End
End
Attribute VB_Name = "frm02010604_3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2021/11/30 改成Form2.0 ; grdDataList改字型=新細明體-ExtB
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/23 日期欄已修改
Option Explicit

'intLastRow上一次反白的Row
'blnOKtoShow決定是否要反白
Dim intLastRow As Integer, blnOKtoShow As Boolean
'intLeaveKind離開時，是0:結束1:回上一畫面
Dim intLeaveKind As Integer
Dim iPrint As Integer, Page As Integer, PLeft(0 To 12) As Integer
Public m_blnFirstShow As Boolean
Public intWhereToGo As Integer '0從frm02010604_1來,1從其他畫面來
Public m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String '本所案號


Private Sub cmdOK_Click(Index As Integer)
Dim intNowRow As Integer
Dim arrCode

    Select Case Index
    Case 0 '確定
        If grdDataList.Rows > 1 Then
            intNowRow = grdDataList.row
            arrCode = Split(grdDataList.TextMatrix(intNowRow, 1), "-")
            frm02010604_2.strCode1 = arrCode(0)
            frm02010604_2.strCode2 = arrCode(1)
            frm02010604_2.strCode3 = arrCode(2)
            frm02010604_2.strCode4 = arrCode(3)
            arrCode = Split(grdDataList.TextMatrix(intNowRow, 4), "-")
            frm02010604_2.strCode5 = arrCode(0)
            frm02010604_2.strCode6 = arrCode(1)
            frm02010604_2.strCode7 = arrCode(2)
            frm02010604_2.strCode8 = arrCode(3)
            frm02010604_2.intChoose = Val(txtChoose)
            frm02010604_2.intWhereToGo = 1
            frm02010604_2.Show
            Me.Hide
        Else
            MsgBox "資料庫無資料 !", vbInformation
        End If
    Case 1 '回前畫面
        intLeaveKind = 1
        Unload Me
    Case 2 '結束
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
        If Replace(grdDataList.TextMatrix(i, 0 + 1), "-", "") = Me.txtCode(0).Text & Me.txtCode(1).Text & IIf(Me.txtCode(0).Text = "TF", Me.txtCode(4).Text, "") & Me.txtCode(2).Text & Me.txtCode(3).Text Then
            grdDataList.TopRow = i
            blnOKtoShow = False
            ShowBar grdDataList, i, Me.grdDataList.Cols - 1
            blnOKtoShow = True
            Exit For
        End If
    Next
End Sub

Private Sub Command2_Click()
On Error GoTo ErrHand
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
      
      strTxt(4) = grdDataList.TextMatrix(i, 7 + 1) & grdDataList.TextMatrix(i, 8 + 1) & _
      grdDataList.TextMatrix(i, 9 + 1) & grdDataList.TextMatrix(i, 10 + 1)
      strTxt(5) = Left(grdDataList.TextMatrix(i, 11 + 1), 10)
      strTxt(6) = grdDataList.TextMatrix(i, 12 + 1)
      strTxt(7) = grdDataList.TextMatrix(i, 13 + 1)
      
      strTxt(8) = grdDataList.TextMatrix(i, 14 + 1)
      strTxt(9) = grdDataList.TextMatrix(i, 15 + 1)
      strTxt(10) = grdDataList.TextMatrix(i, 16 + 1)
      
      For j = 0 To 10
          Printer.CurrentX = PLeft(j)
          Printer.CurrentY = iPrint
          Printer.Print strTxt(j)
      Next
      iPrint = iPrint + 300
      
      If iPrint > 10500 Then
'          Printer.CurrentX = 500
'          Printer.CurrentY = iPrint
'          Printer.Print String(200, "-")
          Printer.NewPage
          Page = Page + 1
          PrintTitle
      End If
   Next
   Printer.EndDoc
   Screen.MousePointer = vbDefault
   MsgBox "列印結束 !", vbInformation
   Exit Sub
ErrHand:
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
   Printer.Print "國內外案件資料維護表"

   iPrint = iPrint + 500
   Printer.Font.Size = 12
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   Printer.CurrentX = 6000
   Printer.CurrentY = iPrint
   
   If frm02010604_1.txtCode(9) <> "" Then
      Printer.Print "國外案工程師：" & frm02010604_1.txtCode(9)
      iPrint = iPrint + 300
   End If
   If frm02010604_1.txtCode(10) <> "" Or frm02010604_1.txtCode(11) <> "" Then
      Printer.Print "國內案發文日：" & ChangeTStringToTDateString(frm02010604_1.txtCode(10)) & " - " & ChangeTStringToTDateString(frm02010604_1.txtCode(11))
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
   Printer.Print "國外案號"
   
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
   Printer.Print "國內案號"
   
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

Private Sub Form_Activate()
Dim varSaveCursor As Variant
Dim ii As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSQL11 As String, strSQL12 As String, strSQL13 As String, strSQL14 As String, strSQL15 As String
Dim strSQL2 As String
Dim StrSQL3 As String
Dim StrSQL4 As String
Dim strSQL51 As String, strSQL52 As String, strSQL53 As String

If m_blnFirstShow = True Then
    Screen.MousePointer = vbHourglass
    grdDataList.MousePointer = vbHourglass
    strSQL11 = "": strSQL12 = "": strSQL13 = "": strSQL14 = "": strSQL15 = ""
    strSQL2 = ""
    StrSQL3 = ""
    StrSQL4 = ""
    strSQL51 = "": strSQL52 = "": strSQL53 = ""
     'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
     Dim midSql As String
     If FMP2open = True Then
        strExc(0) = Replace(FMP2openSQL, "f0.CP", "P1.PA")
        midSql = midSql & strExc(0)
        strExc(0) = Replace(FMP2openSQL, "f0.CP", "P2.PA")
        midSql = midSql & strExc(0)
     End If
    '從分割案維護進入
    If Me.intWhereToGo = 0 Then
        '系統類別
        If frm02010604_1.txtCode(9).Text <> "" Then
            'Patent申請人1~5 (PA26~PA30)
            strSQL11 = strSQL11 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 1) & ") " & midSql
            strSQL12 = strSQL12 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 1) & ") " & midSql
            strSQL13 = strSQL13 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 1) & ") " & midSql
            strSQL14 = strSQL14 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 1) & ") " & midSql
            strSQL15 = strSQL15 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 1) & ") " & midSql
            '商標
            strSQL2 = strSQL2 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 2) & ") "
            StrSQL3 = StrSQL3 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 3) & ") "
            StrSQL4 = StrSQL4 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 4) & ") "
            '服務業務基本資料檔SP08,SP58,SP59
            midSql = Replace(midSql, "P1.PA", "S1.SP")
            midSql = Replace(midSql, "P2.PA", "S2.SP")
            strSQL51 = strSQL51 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 5) & ") " & midSql
            strSQL52 = strSQL52 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 5) & ") " & midSql
            strSQL53 = strSQL53 & " And DC01 In (" & SQLGrpStr(frm02010604_1.txtCode(9).Text, 5) & ") " & midSql
        End If
        '申請人(起)
        If frm02010604_1.txtCode(10).Text <> "" Then
            strSQL11 = strSQL11 & " And P1.PA26>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL12 = strSQL12 & " And P1.PA27>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL13 = strSQL13 & " And P1.PA28>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL14 = strSQL14 & " And P1.PA29>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL15 = strSQL15 & " And P1.PA30>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL2 = strSQL2 & " And T1.TM23>='" & frm02010604_1.txtCode(10).Text & "' "
            StrSQL3 = StrSQL3 & " And L1.LC11>='" & frm02010604_1.txtCode(10).Text & "' "
            StrSQL4 = StrSQL4 & " And H1.HC05>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL51 = strSQL51 & " And S1.SP08>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL52 = strSQL52 & " And S1.SP58>='" & frm02010604_1.txtCode(10).Text & "' "
            strSQL53 = strSQL53 & " And S1.SP59>='" & frm02010604_1.txtCode(10).Text & "' "
        End If
        '申請人(迄)
        If frm02010604_1.txtCode(11).Text <> "" Then
            strSQL11 = strSQL11 & " And P1.PA26<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL12 = strSQL12 & " And P1.PA27<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL13 = strSQL13 & " And P1.PA28<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL14 = strSQL14 & " And P1.PA29<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL15 = strSQL15 & " And P1.PA30<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL2 = strSQL2 & " And T1.TM23<='" & frm02010604_1.txtCode(11).Text & "' "
            StrSQL3 = StrSQL3 & " And L1.LC11<='" & frm02010604_1.txtCode(11).Text & "' "
            StrSQL4 = StrSQL4 & " And H1.HC05<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL51 = strSQL51 & " And S1.SP08<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL52 = strSQL52 & " And S1.SP58<='" & frm02010604_1.txtCode(11).Text & "' "
            strSQL53 = strSQL53 & " And S1.SP59<='" & frm02010604_1.txtCode(11).Text & "' "
        End If
    '從其他畫面進入
    Else
        strSQL11 = strSQL11 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL12 = strSQL12 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL13 = strSQL13 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL14 = strSQL14 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL15 = strSQL15 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL2 = strSQL2 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        StrSQL3 = StrSQL3 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        StrSQL4 = StrSQL4 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL51 = strSQL51 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL52 = strSQL52 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
        strSQL53 = strSQL53 & " And ((DC01='" & m_CP01 & "' And DC02='" & m_CP02 & "' And DC03='" & m_CP03 & "' And DC04='" & m_CP04 & "' ) Or (DC05='" & m_CP01 & "' And DC06='" & m_CP02 & "' And DC07='" & m_CP03 & "' And DC08='" & m_CP04 & "') ) "
    End If
    'Modify By Sindy 2011/2/16 因用SQLDate排序或取MAX或MIN,修改百年蟲問題
'    StrSQLa = "Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, " & SQLDate("P2.PA10") & ") As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
'                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL11
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, " & SQLDate("P2.PA10") & ") As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
'                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL12
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, " & SQLDate("P2.PA10") & ") As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
'                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL13
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, " & SQLDate("P2.PA10") & ") As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
'                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL14
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, " & SQLDate("P2.PA10") & ") As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
'                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL15
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(T1.TM05,Nvl(T1.TM06,T1.TM07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(T2.TM05,Nvl(T2.TM06,T2.TM07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(T2.TM11, Null, Null, " & SQLDate("T2.TM11") & ") As 母案申請日 From DivisionCase, Trademark T1, Trademark T2, Customer C1, Customer C2 Where DC01=T1.TM01(+) And DC02=T1.TM02(+) And DC03=T1.TM03(+) And DC04=T1.TM04(+) And substr(T1.TM23,1,8)=C1.CU01(+) And substr(T1.TM23,9,1)=C1.CU02(+) " & _
'                    " And DC05=T2.TM01(+) And DC06=T2.TM02(+) And DC07=T2.TM03(+) And DC08=T2.TM04(+) And substr(T2.TM23,1,8)=C2.CU01(+) And substr(T2.TM23,9,1)=C2.CU02(+) " & strSQL2
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(L1.LC05,Nvl(L1.LC06,L1.LC07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(L2.LC05,Nvl(L2.LC06,L2.LC07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, ''  As 母案申請日 From DivisionCase, Lawcase L1, Lawcase L2, Customer C1, Customer C2 Where DC01=L1.LC01(+) And DC02=L1.LC02(+) And DC03=L1.LC03(+) And DC04=L1.LC04(+) And substr(L1.LC11,1,8)=C1.CU01(+) And substr(L1.LC11,9,1)=C1.CU02(+) " & _
'                    " And DC05=L2.LC01(+) And DC06=L2.LC02(+) And DC07=L2.LC03(+) And DC08=L2.LC04(+) And substr(L2.LC11,1,8)=C2.CU01(+) And substr(L2.LC11,9,1)=C2.CU02(+) " & StrSQL3
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, H1.HC06 As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, H2.HC06 As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, '' From DivisionCase, Hirecase H1, Hirecase H2, Customer C1, Customer C2 Where DC01=H1.HC01(+) And DC02=H1.HC02(+) And DC03=H1.HC03(+) And DC04=H1.HC04(+) And substr(H1.HC05,1,8)=C1.CU01(+) And substr(H1.HC05,9,1)=C1.CU02(+) " & _
'                    " And DC05=H2.HC01(+) And DC06=H2.HC02(+) And DC07=H2.HC03(+) And DC08=H2.HC04(+) And substr(H2.HC05,1,8)=C2.CU01(+) And substr(H2.HC05,9,1)=C2.CU02(+) " & StrSQL4
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(S1.SP05,Nvl(S1.SP06,S1.SP07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(S2.SP05,Nvl(S2.SP06,S2.SP07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(S2.SP10, Null, Null, " & SQLDate("S2.SP10") & ") As 母案申請日 From DivisionCase, Servicepractice S1, Servicepractice S2, Customer C1, Customer C2 Where DC01=S1.SP01(+) And DC02=S1.SP02(+) And DC03=S1.SP03(+) And DC04=S1.SP04(+) And substr(S1.SP08,1,8)=C1.CU01(+) And substr(S1.SP08,9,1)=C1.CU02(+) " & _
'                    " And DC05=S2.SP01(+) And DC06=S2.SP02(+) And DC07=S2.SP03(+) And DC08=S2.SP04(+) And substr(S2.SP08,1,8)=C2.CU01(+) And substr(S2.SP08,9,1)=C2.CU02(+) " & strSQL51
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(S1.SP05,Nvl(S1.SP06,S1.SP07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(S2.SP05,Nvl(S2.SP06,S2.SP07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(S2.SP10, Null, Null, " & SQLDate("S2.SP10") & ") As 母案申請日 From DivisionCase, Servicepractice S1, Servicepractice S2, Customer C1, Customer C2 Where DC01=S1.SP01(+) And DC02=S1.SP02(+) And DC03=S1.SP03(+) And DC04=S1.SP04(+) And substr(S1.SP08,1,8)=C1.CU01(+) And substr(S1.SP08,9,1)=C1.CU02(+) " & _
'                    " And DC05=S2.SP01(+) And DC06=S2.SP02(+) And DC07=S2.SP03(+) And DC08=S2.SP04(+) And substr(S2.SP08,1,8)=C2.CU01(+) And substr(S2.SP08,9,1)=C2.CU02(+) " & strSQL52
'    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(S1.SP05,Nvl(S1.SP06,S1.SP07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(S2.SP05,Nvl(S2.SP06,S2.SP07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(S2.SP10, Null, Null, " & SQLDate("S2.SP10") & ") As 母案申請日 From DivisionCase, Servicepractice S1, Servicepractice S2, Customer C1, Customer C2 Where DC01=S1.SP01(+) And DC02=S1.SP02(+) And DC03=S1.SP03(+) And DC04=S1.SP04(+) And substr(S1.SP08,1,8)=C1.CU01(+) And substr(S1.SP08,9,1)=C1.CU02(+) " & _
'                    " And DC05=S2.SP01(+) And DC06=S2.SP02(+) And DC07=S2.SP03(+) And DC08=S2.SP04(+) And substr(S2.SP08,1,8)=C2.CU01(+) And substr(S2.SP08,9,1)=C2.CU02(+) " & strSQL53
'    StrSQLa = "Select A.分割案號, A.分割案名, A.分割申請人, A.母案案號, A.母案案名, A.母案申請人, A.母案申請日 From (" & StrSQLa & " ) A Where A.分割案名 Is Not Null "
    StrSQLa = "Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, sqldatet2(P2.PA10)) As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL11
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, sqldatet2(P2.PA10)) As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL12
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, sqldatet2(P2.PA10)) As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL13
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, sqldatet2(P2.PA10)) As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL14
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(P1.PA05,Nvl(P1.PA06,P1.PA07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(P2.PA05,Nvl(P2.PA06,P2.PA07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(P2.PA10, Null, Null, sqldatet2(P2.PA10)) As 母案申請日 From DivisionCase, Patent P1, Patent P2, Customer C1, Customer C2 Where DC01=P1.PA01(+) And DC02=P1.PA02(+) And DC03=P1.PA03(+) And DC04=P1.PA04(+) And substr(P1.PA26,1,8)=C1.CU01(+) And substr(P1.PA26,9,1)=C1.CU02(+) " & _
                    " And DC05=P2.PA01(+) And DC06=P2.PA02(+) And DC07=P2.PA03(+) And DC08=P2.PA04(+) And substr(P2.PA26,1,8)=C2.CU01(+) And substr(P2.PA26,9,1)=C2.CU02(+) " & strSQL15
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(T1.TM05,Nvl(T1.TM06,T1.TM07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(T2.TM05,Nvl(T2.TM06,T2.TM07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(T2.TM11, Null, Null, sqldatet2(T2.TM11)) As 母案申請日 From DivisionCase, Trademark T1, Trademark T2, Customer C1, Customer C2 Where DC01=T1.TM01(+) And DC02=T1.TM02(+) And DC03=T1.TM03(+) And DC04=T1.TM04(+) And substr(T1.TM23,1,8)=C1.CU01(+) And substr(T1.TM23,9,1)=C1.CU02(+) " & _
                    " And DC05=T2.TM01(+) And DC06=T2.TM02(+) And DC07=T2.TM03(+) And DC08=T2.TM04(+) And substr(T2.TM23,1,8)=C2.CU01(+) And substr(T2.TM23,9,1)=C2.CU02(+) " & strSQL2
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(L1.LC05,Nvl(L1.LC06,L1.LC07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(L2.LC05,Nvl(L2.LC06,L2.LC07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, ''  As 母案申請日 From DivisionCase, Lawcase L1, Lawcase L2, Customer C1, Customer C2 Where DC01=L1.LC01(+) And DC02=L1.LC02(+) And DC03=L1.LC03(+) And DC04=L1.LC04(+) And substr(L1.LC11,1,8)=C1.CU01(+) And substr(L1.LC11,9,1)=C1.CU02(+) " & _
                    " And DC05=L2.LC01(+) And DC06=L2.LC02(+) And DC07=L2.LC03(+) And DC08=L2.LC04(+) And substr(L2.LC11,1,8)=C2.CU01(+) And substr(L2.LC11,9,1)=C2.CU02(+) " & StrSQL3
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, H1.HC06 As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, H2.HC06 As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, '' From DivisionCase, Hirecase H1, Hirecase H2, Customer C1, Customer C2 Where DC01=H1.HC01(+) And DC02=H1.HC02(+) And DC03=H1.HC03(+) And DC04=H1.HC04(+) And substr(H1.HC05,1,8)=C1.CU01(+) And substr(H1.HC05,9,1)=C1.CU02(+) " & _
                    " And DC05=H2.HC01(+) And DC06=H2.HC02(+) And DC07=H2.HC03(+) And DC08=H2.HC04(+) And substr(H2.HC05,1,8)=C2.CU01(+) And substr(H2.HC05,9,1)=C2.CU02(+) " & StrSQL4
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(S1.SP05,Nvl(S1.SP06,S1.SP07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(S2.SP05,Nvl(S2.SP06,S2.SP07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(S2.SP10, Null, Null, sqldatet2(S2.SP10)) As 母案申請日 From DivisionCase, Servicepractice S1, Servicepractice S2, Customer C1, Customer C2 Where DC01=S1.SP01(+) And DC02=S1.SP02(+) And DC03=S1.SP03(+) And DC04=S1.SP04(+) And substr(S1.SP08,1,8)=C1.CU01(+) And substr(S1.SP08,9,1)=C1.CU02(+) " & _
                    " And DC05=S2.SP01(+) And DC06=S2.SP02(+) And DC07=S2.SP03(+) And DC08=S2.SP04(+) And substr(S2.SP08,1,8)=C2.CU01(+) And substr(S2.SP08,9,1)=C2.CU02(+) " & strSQL51
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(S1.SP05,Nvl(S1.SP06,S1.SP07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(S2.SP05,Nvl(S2.SP06,S2.SP07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(S2.SP10, Null, Null, sqldatet2(S2.SP10)) As 母案申請日 From DivisionCase, Servicepractice S1, Servicepractice S2, Customer C1, Customer C2 Where DC01=S1.SP01(+) And DC02=S1.SP02(+) And DC03=S1.SP03(+) And DC04=S1.SP04(+) And substr(S1.SP08,1,8)=C1.CU01(+) And substr(S1.SP08,9,1)=C1.CU02(+) " & _
                    " And DC05=S2.SP01(+) And DC06=S2.SP02(+) And DC07=S2.SP03(+) And DC08=S2.SP04(+) And substr(S2.SP08,1,8)=C2.CU01(+) And substr(S2.SP08,9,1)=C2.CU02(+) " & strSQL52
    StrSQLa = StrSQLa & " Union Select DC01||'-'||DC02||'-'||DC03||'-'||DC04 As 分割案號, Nvl(S1.SP05,Nvl(S1.SP06,S1.SP07)) As 分割案名, Nvl(C1.CU04,Decode(C1.CU05, Null,C1.CU06, C1.CU05||' '||C1.CU88||' '||C1.CU89||' '||C1.CU90)) As 分割申請人, DC05||'-'||DC06||'-'||DC07||'-'||DC08 As 母案案號, Nvl(S2.SP05,Nvl(S2.SP06,S2.SP07)) As 母案案名, Nvl(C2.CU04,Decode(C2.CU05, Null,C2.CU06, C2.CU05||' '||C2.CU88||' '||C2.CU89||' '||C2.CU90)) As 母案申請人, Decode(S2.SP10, Null, Null, sqldatet2(S2.SP10)) As 母案申請日 From DivisionCase, Servicepractice S1, Servicepractice S2, Customer C1, Customer C2 Where DC01=S1.SP01(+) And DC02=S1.SP02(+) And DC03=S1.SP03(+) And DC04=S1.SP04(+) And substr(S1.SP08,1,8)=C1.CU01(+) And substr(S1.SP08,9,1)=C1.CU02(+) " & _
                    " And DC05=S2.SP01(+) And DC06=S2.SP02(+) And DC07=S2.SP03(+) And DC08=S2.SP04(+) And substr(S2.SP08,1,8)=C2.CU01(+) And substr(S2.SP08,9,1)=C2.CU02(+) " & strSQL53
    StrSQLa = "Select A.分割案號, A.分割案名, A.分割申請人, A.母案案號, A.母案案名, A.母案申請人, A.母案申請日 From (" & StrSQLa & " ) A Where A.分割案名 Is Not Null "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    Set grdDataList.Recordset = rsA
    grdDataList.Refresh
    SetDataListWidth
    intLastRow = 0
    If grdDataList.Rows > 1 Then
        ShowBar grdDataList, intLastRow, Me.grdDataList.Cols - 1
    End If
    grdDataList.MousePointer = vbDefault
    Screen.MousePointer = vbDefault
    txtChoose.SetFocus
    Me.txtChoose.Text = "5"
    m_blnFirstShow = False
End If
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
m_blnFirstShow = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
If Me.intWhereToGo = 0 Then
    If intLeaveKind = 1 Then
        frm02010604_1.Show
    Else
        Unload frm02010604_1
    End If
End If
intLeaveKind = 0
Set frm02010604_3 = Nothing
End Sub

Private Sub grdDataList_RowColChange()
Dim i As Integer
Dim arrCaseNo

If intLastRow <> grdDataList.row Then
    If blnOKtoShow Then
        blnOKtoShow = False
        ShowBar grdDataList, intLastRow, Me.grdDataList.Cols - 1
        If grdDataList.TextMatrix(grdDataList.row, 1) <> "" Then
            arrCaseNo = Split(grdDataList.TextMatrix(grdDataList.row, 1), "-")
            Me.txtCode(0).Text = arrCaseNo(0)
            If Me.txtCode(0).Text = "TF" Then
                Me.txtCode(1).Text = arrCaseNo(1)
                Me.txtCode(4).Text = arrCaseNo(2)
                Me.txtCode(2).Text = arrCaseNo(3)
                Me.txtCode(3).Text = arrCaseNo(4)
            Else
                Me.txtCode(1).Text = arrCaseNo(1)
                Me.txtCode(2).Text = arrCaseNo(2)
                Me.txtCode(3).Text = arrCaseNo(3)
            End If
        End If
        blnOKtoShow = True
    End If
End If
End Sub

Private Sub SetDataListWidth()
Dim varGridWidth() As Variant
varGridWidth = Array(200, 1500, 1500, 1500, 1500, 1500, 1500, 1000)
SetGridDataListWidth grdDataList, varGridWidth()
Me.grdDataList.TextMatrix(0, 1) = "分割案號"
Me.grdDataList.TextMatrix(0, 2) = "分割案名"
Me.grdDataList.TextMatrix(0, 3) = "分割案申請人"
Me.grdDataList.TextMatrix(0, 4) = "母案案號"
Me.grdDataList.TextMatrix(0, 5) = "母案名稱"
Me.grdDataList.TextMatrix(0, 6) = "母案申請人"
Me.grdDataList.TextMatrix(0, 7) = "母案申請日"
SetDataListVision grdDataList, , True
blnOKtoShow = True
End Sub

Private Sub txtChoose_GotFocus()
txtChoose.SelStart = 0
txtChoose.SelLength = Len(txtChoose)
End Sub

Private Sub txtChoose_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 50 And KeyAscii <> 52 And KeyAscii <> 53 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtChoose_Validate(Cancel As Boolean)
If Val(txtChoose) <> 2 And Val(txtChoose) <> 4 And Val(txtChoose) <> 5 Then
   ShowMsg MsgText(9198)
   txtChoose_GotFocus
   Cancel = True
End If
End Sub

Private Sub txtCode_Change(Index As Integer)
Select Case Index
Case 0 '分割案系統類別
    If Me.txtCode(0).Text = "TF" Then
        Me.txtCode(1).MaxLength = 5
        Me.txtCode(4).Visible = True
        Me.txtCode(4).Enabled = True
        Me.txtCode(4).Text = ""
    Else
        Me.txtCode(1).MaxLength = 6
        Me.txtCode(4).Visible = False
        Me.txtCode(4).Enabled = False
        Me.txtCode(4).Text = ""
    End If
End Select
End Sub

Private Sub txtCode_GotFocus(Index As Integer)
    TextInverse txtCode(Index)
    Command1.Default = True
End Sub

Private Sub txtCode_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub
