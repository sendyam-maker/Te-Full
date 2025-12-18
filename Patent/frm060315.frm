VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm060315 
   BorderStyle     =   1  '單線固定
   Caption         =   "翻譯未分案明細表"
   ClientHeight    =   5760
   ClientLeft      =   2475
   ClientTop       =   4515
   ClientWidth     =   9375
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   9375
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   8310
      TabIndex        =   12
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6300
      TabIndex        =   11
      Top             =   90
      Width           =   756
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印(&P)"
      Height          =   400
      Left            =   7110
      TabIndex        =   10
      Top             =   90
      Width           =   1155
   End
   Begin VB.TextBox txtPA08 
      Height          =   264
      Left            =   1305
      MaxLength       =   1
      TabIndex        =   4
      Top             =   720
      Width           =   450
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   0
      Top             =   90
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   2265
      MaxLength       =   7
      TabIndex        =   1
      Top             =   90
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   2265
      MaxLength       =   7
      TabIndex        =   3
      Top             =   405
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1305
      MaxLength       =   7
      TabIndex        =   2
      Top             =   405
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
      Height          =   4575
      Left            =   180
      TabIndex        =   9
      Top             =   1080
      Width           =   9090
      _ExtentX        =   16034
      _ExtentY        =   8070
      _Version        =   393216
      Cols            =   15
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      FormatString    =   $"frm060315.frx":0000
      _NumberOfBands  =   1
      _Band(0).Cols   =   15
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "( 1 發明 2 新型 3 設計 )"
      Height          =   180
      Index           =   3
      Left            =   1800
      TabIndex        =   8
      Top             =   780
      Width           =   1785
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "種類："
      Height          =   180
      Index           =   1
      Left            =   750
      TabIndex        =   7
      Top             =   780
      Width           =   540
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "申請案發文日："
      Height          =   180
      Index           =   2
      Left            =   30
      TabIndex        =   6
      Top             =   150
      Width           =   1260
   End
   Begin VB.Line Line3 
      X1              =   2190
      X2              =   2295
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line Line1 
      X1              =   2190
      X2              =   2295
      Y1              =   555
      Y2              =   555
   End
   Begin VB.Label Label1 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      Caption         =   "收文日："
      Height          =   180
      Index           =   0
      Left            =   570
      TabIndex        =   5
      Top             =   465
      Width           =   720
   End
End
Attribute VB_Name = "frm060315"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
'整理 by Morgan 2005/9/13
Option Explicit

Dim PLeft(0 To 12) As Integer, iPrint As Integer, iPage As Integer, strTemp(1 To 11) As String
Private Const ciTitleFontSize = 22, ciFontSize = 12
Private Const ciStartX = 500, ciStartY = 500
Dim lngPageHeight As Long, lngPageWidth As Long, lngLineHeight As Long


Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdPrint_Click()
   If txtValidate = True Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Process 1
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub cmdQuery_Click()
   If txtValidate = True Then
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      Process
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   GetPleft
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm060315 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'p_iMode:0=查詢,1=列印
Sub Process(Optional ByVal p_iMode As Integer = "0")
   Dim stCon As String
   
   ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/13 清除查詢印表記錄檔欄位
   stCon = ""
   If Trim(txt1(0)) <> "" Then
      stCon = stCon & " AND A.CP05>=" & TransDate(txt1(0), 2)
   End If
   If Trim(txt1(1)) <> "" Then
      stCon = stCon & " AND A.CP05<=" & TransDate(txt1(1), 2)
   End If
   If Trim(txt1(0)) <> "" Or Trim(txt1(1)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/12/13
   End If
   
   If Trim(txt1(2)) <> "" Then
      stCon = stCon & " AND B.CP27>=" & TransDate(txt1(2), 2)
   End If
   If Trim(txt1(3)) <> "" Then
      stCon = stCon & " AND B.CP27<=" & TransDate(txt1(3), 2)
   End If
   If Trim(txt1(2)) <> "" Or Trim(txt1(3)) <> "" Then
      pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/12/13
   End If
   
   If Trim(txtPA08) <> "" Then
      stCon = stCon & " AND PA08='" & txtPA08 & "'"
      pub_QL05 = pub_QL05 & ";" & Label1(1) & txtPA08 & Label1(3) 'Add By Sindy 2010/12/13
   End If
   
   '2009/8/21 MODIFY BY SONIA 加閉卷資料不要印的條件
   strSql = "SELECT B.CP27-19110000 R01" & _
      ", A.CP05-19110000 R02, PA10-19110000 R03, A.CP09 R04, A.CP01||'-'||A.CP02||'-'||A.CP03||'-'||A.CP04 R05" & _
      ", PA05 R06,PTM03 R07, A.CP06-19110000 R08, A.CP07-19110000 R09,ST02 R10, SUBSTR(A.CP64,1,15) R11" & _
      " FROM CASEPROGRESS B,CASEPROGRESS A,PATENT,FAGENT,NATION,STAFF,PATENTTRADEMARKMAP" & _
      " WHERE B.CP01='FCP' AND B.CP27>0 AND B.CP57 IS NULL AND B.CP31='Y' AND B.CP09<'C'" & _
      " AND A.CP01(+)=B.CP01 AND A.CP02(+)=B.CP02 AND A.CP03(+)=B.CP03 AND A.CP04(+)=B.CP04" & _
      " AND A.CP10='201' AND A.CP14 IS NULL AND A.CP27 IS NULL AND A.CP57 IS NULL" & stCon & _
      "" & _
      " AND PA01(+)=B.CP01 AND PA02(+)=B.CP02 AND PA03(+)=B.CP03 AND PA04(+)=B.CP04 AND PA57 IS NULL" & _
      " AND FA01(+)=SUBSTR(PA75,1,8) AND FA02(+)=SUBSTR(PA75,9,1)" & _
      " AND NA01(+)=FA10 AND ST01(+)=NA16" & _
      " AND PTM01(+)='1' AND PTM02(+)=PA08" & _
      " ORDER BY 1,2,3,4"
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/13
         Set GrdDataList.Recordset = adoRecordset.Clone
         GrdDataList.FormatString = GrdDataList.FormatString
         If p_iMode = 1 Then
            PrintData adoRecordset
         End If
      Else
         InsertQueryLog (0) 'Add By Sindy 2010/12/13
         ShowNoData
      End If
   End With
End Sub

Sub PrintData(ByRef p_Rst As ADODB.Recordset)
   Dim iOrientation As Integer, iRow As Integer, iCol As Integer
   
   iOrientation = Printer.Orientation
   Printer.Orientation = 2
   lngPageHeight = Printer.ScaleHeight
   lngPageWidth = Printer.ScaleWidth
   lngLineHeight = 300
   
   With p_Rst
      iPage = 1
      PrintPageHeader
      PrintPageHeader1
      .MoveFirst
      Do While Not .EOF
         For iCol = 1 To 11
            Select Case iCol
               Case 6
                  strTemp(iCol) = Left("" & .Fields(iCol - 1), 7)
               Case 1, 2, 3, 8, 9
                  strTemp(iCol) = Format(TransDate("" & .Fields(iCol - 1), 1), "###/##/##")
               Case 11
                  strTemp(iCol) = Left("" & .Fields(iCol - 1), 3) & IIf(Len("" & .Fields(iCol - 1)) > 3, "...", "")
               Case Else
                  strTemp(iCol) = "" & .Fields(iCol - 1)
            End Select
         Next
         PrintDetail
         .MoveNext
      Loop
      Call PrintReportFooter(.RecordCount)
   End With
   Printer.EndDoc
   MsgBox "列印完成！"
   Printer.Orientation = iOrientation
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) <> "" Then
      If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
         Cancel = True
         Me.txt1(Index).SetFocus
         txt1_GotFocus Index
      End If
   End If
End Sub

Private Function txtValidate() As Boolean
   Dim bolCancel As Boolean, ii As Integer
   
   If Trim(txt1(2)) & Trim(txt1(3)) = "" Then
      MsgBox Label1(2) & "區間不可空白!!", vbExclamation, "USER 輸入錯誤"
      If Trim(txt1(0)) = "" Then
         txt1(2).SetFocus
      Else
         txt1(3).SetFocus
      End If
      Exit Function
   End If
   
   For ii = 0 To 3
      txt1_Validate ii, bolCancel
      If bolCancel = True Then
         Me.txt1(ii).SetFocus
         txt1_GotFocus ii
         Exit Function
      End If
   Next
            
   If Val(Me.txt1(0).Text) > Val(Me.txt1(1).Text) Then
      MsgBox "收文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
      Me.txt1(0).SetFocus
      txt1_GotFocus 0
      Exit Function
   End If
         
   If Me.txt1(2).Text <> "" And Me.txt1(3).Text <> "" Then
      If Val(Me.txt1(2).Text) > Val(Me.txt1(3).Text) Then
         MsgBox "發文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
         Me.txt1(2).SetFocus
         txt1_GotFocus 2
         Exit Function
      End If
   End If
        
   txtValidate = True
End Function

Sub GetPleft()
    Erase PLeft
    PLeft(0) = 500
    PLeft(1) = 500
    '申請發文日(1600)
    PLeft(2) = PLeft(1) + 1600
    '收文日(1200)
    PLeft(3) = PLeft(2) + 1200
    '申請日(1200)
    PLeft(4) = PLeft(3) + 1200
    '總收文號(1400)
    PLeft(5) = PLeft(4) + 1400
    '本所案號(2000)
    PLeft(6) = PLeft(5) + 2000
    '案件名稱(2100)
    PLeft(7) = PLeft(6) + 2100
    '種類(900)
    PLeft(8) = PLeft(7) + 900
    '本所期限(1200)
    PLeft(9) = PLeft(8) + 1200
    '法定期限(1200)
    PLeft(10) = PLeft(9) + 1200
    '管制人(1200)
    PLeft(11) = PLeft(10) + 1200
    '備註
End Sub

Private Sub PrintNewLine(Optional ByVal bolSubtotal As Boolean = True, Optional ByVal iExtraLines As Integer = 3)
   iPrint = iPrint + lngLineHeight
   If iPrint >= (lngPageHeight - iExtraLines * lngLineHeight) Then
      Printer.CurrentX = ciStartX
      Printer.CurrentY = iPrint
      Printer.Print String(125, "-")
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
      Printer.CurrentX = PLeft(iCol)
      Printer.CurrentY = iPrint
      Printer.Print strTemp(iCol)
    Next
End Sub

Sub PrintPageHeader()
   Dim strPTmp As String
    iPrint = ciStartY
    Printer.FontName = "細明體"
    Printer.Font.Size = ciTitleFontSize
    Printer.Font.Bold = True
    Printer.Font.Underline = True
    strPTmp = "FCP翻譯未分案明細表"
    Printer.CurrentX = (lngPageWidth - Printer.TextWidth(strPTmp)) / 2
    Printer.CurrentY = iPrint
    Printer.Print strPTmp
    iPrint = iPrint + 500
    Printer.Font.Size = ciFontSize
    Printer.Font.Bold = False
    Printer.Font.Underline = False
    
    strPTmp = "申請案發文日：" & txt1(2) & " - " & txt1(3)
    Printer.CurrentX = (lngPageWidth - Printer.TextWidth(Space(30))) / 2
    Printer.CurrentY = iPrint
    Printer.Print strPTmp
    
    If Trim(txt1(0) & txt1(1)) <> "" Then
      PrintNewLine
      strPTmp = "　　　收文日：" & txt1(0) & " - " & txt1(1)
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(Space(30))) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
    End If
    
    If Trim(txtPA08) <> "" Then
      PrintNewLine
      strPTmp = "　　　　種類：" & txtPA08
      Printer.CurrentX = (lngPageWidth - Printer.TextWidth(Space(30))) / 2
      Printer.CurrentY = iPrint
      Printer.Print strPTmp
    End If
    
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
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(125, "-")
End Sub

Sub PrintPageHeader1()
    Call PrintNewLine(False, 1)
    
    Printer.CurrentX = PLeft(1)
    Printer.CurrentY = iPrint
    Printer.Print "申請案發文日"
    
    Printer.CurrentX = PLeft(2)
    Printer.CurrentY = iPrint
    Printer.Print "收文日"
    
    Printer.CurrentX = PLeft(3)
    Printer.CurrentY = iPrint
    Printer.Print "申請日"
    
    Printer.CurrentX = PLeft(4)
    Printer.CurrentY = iPrint
    Printer.Print "總收文號"
    
    Printer.CurrentX = PLeft(5)
    Printer.CurrentY = iPrint
    Printer.Print "本所案號"
    
    Printer.CurrentX = PLeft(6)
    Printer.CurrentY = iPrint
    Printer.Print "案件名稱"
    
    Printer.CurrentX = PLeft(7)
    Printer.CurrentY = iPrint
    Printer.Print "種類"
    
    Printer.CurrentX = PLeft(8)
    Printer.CurrentY = iPrint
    Printer.Print "本所期限"
    
    Printer.CurrentX = PLeft(9)
    Printer.CurrentY = iPrint
    Printer.Print "法定期限"
    
    Printer.CurrentX = PLeft(10)
    Printer.CurrentY = iPrint
    Printer.Print "管制人"
    
    Printer.CurrentX = PLeft(11)
    Printer.CurrentY = iPrint
    Printer.Print "備註"
    
    PrintNewLine
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(125, "-")
End Sub

'列印表尾
Private Sub PrintReportFooter(Optional ByVal iRecCount As Integer = 0)
    iPrint = iPrint + lngLineHeight
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print String(125, "-")
    
    iPrint = iPrint + lngLineHeight
    Printer.CurrentX = ciStartX
    Printer.CurrentY = iPrint
    Printer.Print "合計： " & iRecCount & " 筆"
    Printer.EndDoc
End Sub

Private Sub txtPA08_GotFocus()
   TextInverse txtPA08
End Sub

Private Sub txtPA08_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") And KeyAscii <> Asc("3") Then
      KeyAscii = 0
      Beep
   End If
End Sub
