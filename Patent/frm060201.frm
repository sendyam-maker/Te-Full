VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060201 
   BorderStyle     =   1  '單線固定
   Caption         =   "核稿期限管制查詢/列印"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9375
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   7065
      TabIndex        =   4
      Top             =   60
      Width           =   1200
   End
   Begin VB.TextBox txtDate 
      Height          =   264
      Left            =   1500
      MaxLength       =   7
      TabIndex        =   0
      Top             =   210
      Width           =   1215
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6255
      TabIndex        =   1
      Top             =   60
      Width           =   756
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8325
      TabIndex        =   2
      Top             =   60
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4845
      Left            =   135
      TabIndex        =   5
      Top             =   570
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   8546
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   "核稿期限|承辦期限|本所案號　　　　|核稿人|承辦人|案件名稱　　　|申請日　|完稿日　|是否會稿|本所期限|法定期限|種類　"
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
      _Band(0).Cols   =   12
   End
   Begin VB.Label lblMemo 
      AutoSize        =   -1  'True
      Caption         =   "* 為已延期, ** 為未完稿無核稿期限"
      Height          =   180
      Left            =   180
      TabIndex        =   6
      Top             =   5520
      Width           =   2790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "管制期限止日："
      Height          =   180
      Left            =   180
      TabIndex        =   3
      Top             =   240
      Width           =   1260
   End
End
Attribute VB_Name = "frm060201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/12/16 Form2.0已修改 (Printer列印未改)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim PLeft(0 To 12) As Integer, iPrint As Integer, iPage As Integer, strTemp(1 To 12) As String
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long


Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub DoPrint()
   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
 
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   With grdDataList
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      For iRow = 1 To .Rows - 1
          For iCol = LBound(strTemp) To UBound(strTemp)
            If iCol = 6 Then
               strTemp(iCol) = Left(.TextMatrix(iRow, iCol - 1), 7)
            Else
               strTemp(iCol) = .TextMatrix(iRow, iCol - 1)
            End If
          Next
          PrintDetail
      Next
      Call PrintReportFooter(.Rows - 1)
      Printer.EndDoc
      MsgBox "列印完成！"
   End With
   Printer.Orientation = iOrientation
End Sub

Private Function TxtValidate() As Boolean

   Dim bolCancel As Boolean
   
   If txtDate = "" Then
      MsgBox Label1 & "不可空白！", vbExclamation
      If txtDate.Enabled = True Then
         txtDate.SetFocus
      End If
      GoTo flgFail
   Else
      bolCancel = False
      Call txtDate_Validate(bolCancel)
      If bolCancel Then GoTo flgFail
   End If
      
   TxtValidate = True
   
flgFail:

End Function

Sub GetPleft()

    Erase PLeft
    PLeft(0) = 500
    PLeft(1) = 500
    '核稿期限(1200)
    PLeft(2) = PLeft(1) + 1200
    '承辦期限(1200)
    PLeft(3) = PLeft(2) + 1200
    '本所案號(2000)
    PLeft(4) = PLeft(3) + 2000
    '核稿人(1050)
    PLeft(5) = PLeft(4) + 1050
    '承辦人(1050)
    PLeft(6) = PLeft(5) + 1050
    '案件名稱(2100)
    PLeft(7) = PLeft(6) + 2100
    '申請日(1200)
    PLeft(8) = PLeft(7) + 1200
    '完稿日(1200)
    PLeft(9) = PLeft(8) + 1200
    '是否會稿(1200)
    PLeft(10) = PLeft(9) + 1200
    '本所期限(1200)
    PLeft(11) = PLeft(10) + 1200
    '法定期限(1200)
    PLeft(12) = PLeft(11) + 1200
    '種類(900)
    
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

Sub PrintDetail()

    Dim iCol As Integer

    PrintNewLine
    For iCol = LBound(strTemp) To UBound(strTemp)
      If iCol = 3 Then
        Printer.CurrentX = PLeft(iCol + 1) - Printer.TextWidth(strTemp(iCol)) - 100
        Printer.CurrentY = iPrint
        Printer.Print strTemp(iCol)
      Else
        Printer.CurrentX = PLeft(iCol)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(iCol)
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
    strPTmp = "FCP核稿期限管制表"
    Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
    Printer.CurrentY = iPrint
    Printer.Print strPTmp
    iPrint = iPrint + 500
    Printer.Font.Size = ciFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    
    strPTmp = "管制期限止日：" & txtDate
    Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
    Printer.CurrentY = iPrint
    Printer.Print strPTmp
    
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "列印人：" & strUserName
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
    PrintNewLine
    Printer.CurrentX = 13000
    Printer.CurrentY = iPrint
    Printer.Print "頁    次：" & str(iPage)
    PrintNewLine
    Printer.CurrentX = 6000
    Printer.CurrentY = iPrint
    Printer.Print lblMemo.Caption
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(200, "-")
End Sub

Sub PrintPageHeader1()

    Call PrintNewLine(False, 1)
    
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "核稿期限"
    
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "承辦期限"
    
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = iPrint
    Printer.Print "本所案號"
    
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = iPrint
    Printer.Print "核稿人"
    
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = iPrint
    Printer.Print "承辦人"
    
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "案件名稱"
    
    Printer.CurrentX = PLeft(7)
    Printer.CurrentY = iPrint
    Printer.Print "申請日"
    
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = iPrint
    Printer.Print "完稿日"
    
    Printer.CurrentX = PLeft(9)
    Printer.CurrentY = iPrint
    Printer.Print "是否會稿"
    
    Printer.CurrentX = PLeft(10)
    Printer.CurrentY = iPrint
    Printer.Print "本所期限"
    
    Printer.CurrentX = PLeft(11)
    Printer.CurrentY = iPrint
    Printer.Print "法定期限"
    
    Printer.CurrentX = PLeft(12)
    Printer.CurrentY = iPrint
    Printer.Print "種類"
    
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
   If grdDataList.TextMatrix(1, 2) <> "" Then
      DoPrint
   End If
End Sub

Private Sub doQuery()
Dim stCon As String, stTmp As String
   
On Error GoTo flgErr
    
    ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
    
    If txtDate <> "" Then
      stTmp = TransDate(txtDate, 2)
      stCon = " AND A.CP48<=" & stTmp & " AND ( EP08<=" & stTmp & " OR EP08 IS NULL )"
      pub_QL05 = pub_QL05 & ";" & Label1 & txtDate 'Add By Sindy 2010/12/7
    End If
    'Modify by Morgan 2010/8/13 百年蟲
    strSql = "Select DECODE(EP08,NULL,DECODE(EP09,NULL,'   **',''),substrb(' '||sqldatet(EP08),-9)) R01" & _
      ", substrb(' '||sqldatet(A.CP48),-9) R02" & _
      ", DECODE(INSTR('201,210',A.CP10),0,NULL,DECODE(F.CP01,NULL,NULL,'*'))||A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04 R03" & _
      ", S1.ST02 R04, S2.ST02 R05" & _
      ", PA05 R06" & _
      ", substrb(' '||sqldatet(PA10),-9) R07" & _
      ", substrb(' '||sqldatet(EP09),-9) R08" & _
      ", EP34 R09" & _
      ", substrb(' '||sqldatet(A.CP06),-9) R10" & _
      ", substrb(' '||sqldatet(A.CP07),-9) R11" & _
      ", PTM03 R12" & _
      " From CASEPROGRESS A, ENGINEERPROGRESS B, PATENT C, PATENTTRADEMARKMAP E, STAFF S1, STAFF S2,CASEPROGRESS F" & _
      " Where A.CP27 IS NULL AND A.CP57 IS NULL" & stCon & _
      " AND A.CP01='FCP' AND A.CP10 IN ('201','209','210') AND A.CP14 IS NOT NULL" & _
      " AND EP02(+)=A.CP09 AND EP07 IS NULL" & _
      " AND PA01(+)=A.CP01 AND PA02(+)=A.CP02 AND PA03(+)=A.CP03 AND PA04(+)=A.CP04 AND PA57 IS NULL " & _
      " AND PTM01(+)='1' AND PTM02(+)=PA08" & _
      " AND S1.ST01(+)=EP04 AND S2.ST01(+)=A.CP14" & _
      " AND F.CP43(+)=A.CP09 AND F.CP10(+)='404'" & _
      " ORDER BY EP08,2,3,12"

    CheckOC
    
    With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/7
         Set grdDataList.Recordset = adoRecordset.Clone
         grdDataList.FormatString = grdDataList.FormatString
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/7
         MsgBox "查無資料！", vbInformation
      End If
   End With
   
flgErr:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
    
End Sub

Private Sub cmdQuery_Click()
    Screen.MousePointer = vbHourglass
    If TxtValidate Then doQuery
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    GetPleft
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060201 = Nothing
End Sub

Private Sub txtDate_GotFocus()
   TextInverse txtDate
End Sub

Private Sub txtDate_Validate(Cancel As Boolean)
   If txtDate <> "" Then
      If Not ChkDate(txtDate) Then
        Cancel = True
      End If
   End If
End Sub
