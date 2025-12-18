VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060202 
   BorderStyle     =   1  '單線固定
   Caption         =   "員工翻譯費率查詢/列印"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7350
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7350
   Begin VB.TextBox txtState 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   3
      Top             =   930
      Width           =   525
   End
   Begin VB.TextBox txtDept 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   2
      Top             =   600
      Width           =   525
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   1
      Left            =   2565
      MaxLength       =   6
      TabIndex        =   1
      Top             =   235
      Width           =   1080
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   5130
      TabIndex        =   5
      Top             =   60
      Width           =   1200
   End
   Begin VB.TextBox txtNo 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   264
      Index           =   0
      Left            =   1305
      MaxLength       =   6
      TabIndex        =   0
      Top             =   235
      Width           =   1080
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   4050
      TabIndex        =   4
      Top             =   60
      Width           =   1020
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   6390
      TabIndex        =   6
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4335
      Left            =   135
      TabIndex        =   8
      Top             =   1260
      Width           =   7065
      _ExtentX        =   12462
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "員工代號|姓名　　　|英文翻譯費率|日文翻譯費率|中文打字費率|身分|狀態"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   7
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "是否含離職：            (Y: 含離職)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   960
      Width           =   2640
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "費率計算單位 : NT$/千字"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   4860
      TabIndex        =   10
      Top             =   990
      Width           =   2070
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "身分：　                     (1:內翻 2:外翻)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   630
      Width           =   2940
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號：                             －"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.75
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   270
      Width           =   2475
   End
End
Attribute VB_Name = "frm060202"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改 (Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim PLeft() As Integer, iPrint As Integer, iPage As Integer
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long
Dim bolBarShow As Boolean


Private Sub Form_Activate()
   bolBarShow = Forms(0).StatusBar1.Visible
   Forms(0).StatusBar1.Visible = True
End Sub

Private Sub Form_Deactivate()
   Forms(0).StatusBar1.Visible = bolBarShow
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub DoPrint()

   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   Dim strTemp() As String
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 1
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      GetPleft
      intI = .Cols
      ReDim strTemp(1 To intI)
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
         For iCol = LBound(strTemp) To UBound(strTemp)
            If iCol > 2 And iCol < 6 Then
               strTemp(iCol) = Format(.TextMatrix(iRow, iCol - 1), "#,##0")
            Else
               strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
            End If
         Next
         PrintDetail strTemp
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
   
   
End Sub

Sub GetPleft()
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   intI = grdDataList.Cols
   ReDim PLeft(1 To intI)
   PLeft(1) = ciStartX
   For intI = 2 To grdDataList.Cols
      PLeft(intI) = PLeft(intI - 1) + Printer.TextWidth(grdDataList.TextMatrix(0, intI - 2)) + 150
   Next
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 2)

   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.Print String(200, "-")
      iPage = iPage + 1
      Printer.NewPage
      PrintPageHeader
      If bolSubtotal Then
         PrintPageHeader1
         iPrint = iPrint + lngLineHeight
      End If
   End If
    
End Sub

Sub PrintDetail(strData() As String)
    Dim iCol As Integer
    PrintNewLine
    For iCol = LBound(strData) To UBound(strData)
      If iCol > 2 And iCol < 6 Then
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strData(iCol)) - 150
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      Else
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strData(iCol)
      End If
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
   iPrint = ciStartY
   Printer.FontName = "細明體"
   Printer.Font.Size = ciTitleFontSize
   Printer.Font.Bold = True
   Printer.Font.Underline = True
   strPTmp = "員工翻譯費率表"
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   iPrint = iPrint + 500
   Printer.Font.Size = ciFontSize
   Printer.Font.Bold = False
   Printer.Font.Underline = False
   
   strPTmp = "員工代號："
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp & txtNo(0) & " － " & txtNo(1)
    
   If txtDept = "1" Then
     strPTmp = "身份：內翻"
   ElseIf txtDept = "2" Then
     strPTmp = "身份：外翻"
   Else
     strPTmp = "身份別：全部"
   End If
   PrintNewLine
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth("身份：")) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
   
   If txtState = "Y" Then
      strPTmp = "是否含離職：含"
   Else
      strPTmp = "是否含離職：不含"
   End If
   
   PrintNewLine
   Printer.CurrentX = (lngPageWidth - Printer.TextWidth("是否含離職：")) / 2
   Printer.CurrentY = iPrint
   Printer.Print strPTmp
    
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print "列印人：" & strUserName
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
   PrintNewLine
   Printer.CurrentX = lngPageWidth - Printer.TextWidth(String(12, "　"))
   Printer.CurrentY = iPrint
   Printer.Print "頁    次：" & str(iPage)
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print lblMemo.Caption
   PrintNewLine
   Printer.CurrentX = ciStartX
   Printer.CurrentY = iPrint
   Printer.Print String(200, "-")
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    For intI = 1 To grdDataList.Cols
      Printer.CurrentX = PLeft(intI)
      Printer.CurrentY = iPrint
      Printer.Print grdDataList.TextMatrix(0, intI - 1)
    Next
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    
End Sub
'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)

    Call PrintNewLine(True, 1)
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Private Sub cmdPrint_Click()
   If Not grdDataList.Recordset Is Nothing Then
      If grdDataList.Recordset.RecordCount > 0 Then
         DoPrint
      End If
   End If
End Sub

Private Sub doQuery()
Dim stCon As String
   
On Error GoTo flgErr
    
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
    stCon = ""
    If txtNo(0) <> "" Then
      stCon = stCon & " and spr01>='" & txtNo(0) & "'"
    End If
    If txtNo(1) <> "" Then
      stCon = stCon & " and spr01<='" & txtNo(1) & "'"
    End If
    If txtNo(0) <> "" Or txtNo(1) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Left(Label1, 5) & txtNo(0) & "-" & txtNo(1) 'Add By Sindy 2010/12/7
    End If
    '有對應到所內員工編號的外譯編號為"內翻"其餘皆為"外翻"
    If txtDept = "1" Then
      stCon = stCon & " and D.st04='1'"
      pub_QL05 = pub_QL05 & ";" & Left(Label2, 3) & "1:內翻" 'Add By Sindy 2010/12/7
    ElseIf txtDept = "2" Then
      stCon = stCon & " and ( D.st04 is null or D.st04<>'1')"
      pub_QL05 = pub_QL05 & ";" & Left(Label2, 3) & "2:外翻" 'Add By Sindy 2010/12/7
    End If
    
    If txtState <> "Y" Then
      stCon = stCon & " and B.st04='1'"
   Else
      pub_QL05 = pub_QL05 & ";" & Left(Label3, 6) & txtState 'Add By Sindy 2010/12/7
    End If
    
    strSql = "Select SPR01,B.ST02,SPR02,SPR03,SPR04,DECODE(D.st04,'1','內翻','外翻') C06,DECODE(B.ST04,'1','在職','離職') C07" & _
      " From STAFF_PAYRATE A, STAFF B,staff_idmap C,staff D" & _
      " Where B.ST01=SPR01 and sim02(+)=spr01 and D.st01(+)=sim01" & stCon & _
      " ORDER BY 6,1"
    
    CheckOC
    
    With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      Set grdDataList.Recordset = adoRecordset.Clone
      grdDataList.FormatString = grdDataList.FormatString
      RecordShow
      If .RecordCount = 0 Then
         InsertQueryLog (0) 'Add By Sindy 2010/12/7
         MsgBox "查無資料！", vbInformation
      Else
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
      End If
   End With
   
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
    
End Sub

Private Sub cmdQuery_Click()
    Screen.MousePointer = vbHourglass
    doQuery
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060202 = Nothing
End Sub

'*************************************************
'  顯示筆數
'
'*************************************************
Private Sub RecordShow()
   Forms(0).StatusBar1.Panels(2).Text = grdDataList.Recordset.RecordCount
End Sub

Private Sub txtDept_GotFocus()
   TextInverse txtDept
End Sub

Private Sub txtDept_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
      Beep
      KeyAscii = 0
   End If
End Sub

Private Sub txtNo_GotFocus(Index As Integer)
   If Index = 1 Then
      If txtNo(0) <> "" And txtNo(1) = "" Then
         txtNo(1) = txtNo(0)
      End If
   End If
   TextInverse txtNo(Index)
End Sub

Private Sub txtNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtState_GotFocus()
   TextInverse txtState
End Sub

Private Sub txtState_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
      Beep
      KeyAscii = 0
   End If
End Sub
