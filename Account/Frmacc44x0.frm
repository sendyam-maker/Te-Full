VERSION 5.00
Begin VB.Form Frmacc44x0 
   AutoRedraw      =   -1  'True
   Caption         =   "年度扣繳檢核(抬頭及信箱)"
   ClientHeight    =   4030
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5560
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4030
   ScaleWidth      =   5560
   Begin VB.TextBox txtYear 
      Height          =   290
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   6
      Top             =   240
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Caption         =   "檢核清單"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2330
      Left            =   330
      TabIndex        =   4
      Top             =   780
      Width           =   4880
      Begin VB.CheckBox Check3 
         Caption         =   "有代表信箱，但無財務、會計師信箱"
         Height          =   255
         Left            =   450
         TabIndex        =   2
         Top             =   1890
         Width           =   3500
      End
      Begin VB.CheckBox Check2 
         Caption         =   "代表、財務、會計師信箱均無建立"
         Height          =   255
         Left            =   450
         TabIndex        =   1
         Top             =   1580
         Width           =   3500
      End
      Begin VB.CheckBox Check1 
         Caption         =   "扣繳收據抬頭檢查清單"
         Height          =   255
         Left            =   450
         TabIndex        =   0
         Top             =   960
         Width           =   3500
      End
      Begin VB.Label Label4 
         Caption         =   "今年有開立收據 + 扣繳年度(有扣繳)"
         ForeColor       =   &H00FF0000&
         Height          =   250
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   3700
      End
      Begin VB.Label Label3 
         Caption         =   $"Frmacc44x0.frx":0000
         ForeColor       =   &H00FF0000&
         Height          =   610
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   4540
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel(&E)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1620
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   3510
      Width           =   2115
   End
   Begin VB.Label Label2 
      Caption         =   "扣繳年度："
      Height          =   290
      Left            =   1080
      TabIndex        =   7
      Top             =   270
      Width           =   920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "注意: 依收據抬頭抓取相關資料會花費一些時間, 請稍待片刻~"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   220
      Left            =   180
      TabIndex        =   5
      Top             =   3240
      Width           =   5230
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc44x0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/10/12 Form2.0已檢查 (無需修改的物件)
'Create By Sindy 2017/11/6
Option Explicit

Dim adoquery As New ADODB.Recordset
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim intCounter As Integer
Dim lngPageNo As Long '頁數


Private Sub Check1_Click()
   If Check1.Value = 1 Then
      Check2.Value = 0
      Check3.Value = 0
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = 1 Then
      Check1.Value = 0
   End If
End Sub

Private Sub Check3_Click()
   If Check3.Value = 1 Then
      Check1.Value = 0
   End If
End Sub

Private Sub Command1_Click()
Dim strErr As String, StrOk As String
   
   strErr = "": StrOk = ""
   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   
   Command1.Enabled = False
   'Add By Sindy 2025/6/26 使用共用函數,抓資料寫入暫存檔中,供下列程式做使用
   If Me.Command1.Tag = IIf(Check1.Value = 1, "1", "2") & "Y" Then
      If MsgBox("需要再重新解析資料至暫存檔(供報表使用)嗎？" & vbCrLf & vbCrLf & _
                "是: 重新讀取資料(需花費一些時間, 請稍待片刻~)" & vbCrLf & _
                "否: 直接產生報表", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
         Me.Command1.Tag = ""
      End If
   Else
      Me.Command1.Tag = ""
   End If
   If Me.Command1.Tag = "" Then
      Call PUB_Frmacc44x0(Me.Name, IIf(Check1.Value = 1, "1", "2"), txtYear, strSrvDate(2))
      Me.Command1.Tag = IIf(Check1.Value = 1, "1", "2") & "Y"
   End If
   
   'Modify By Sindy 2021/10/8
   If Check1.Value = 1 Then
      If Process = False Then
         strErr = strErr & "扣繳收據抬頭檢查清單;"
      Else
         StrOk = "Y"
      End If
   End If
   If Check2.Value = 1 Then
      If Process2 = False Then
         strErr = strErr & "代表、財務、會計師信箱均無建立;"
      Else
         StrOk = "Y"
      End If
   End If
   If Check3.Value = 1 Then
      If Process3 = False Then
         strErr = strErr & "有代表信箱，但無財務、會計師信箱;"
      Else
         StrOk = "Y"
      End If
   End If
   '2021/10/8 END
   
   If UCase(strErr) <> "" Then MsgBox strErr & " 無資料，可供列印！"
   If UCase(StrOk) = "Y" Then MsgBox "資料產生完畢！"
   
   Command1.Enabled = True
End Sub

Private Function Process() As Boolean
  
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   
'   '抬頭4個字(含)以上但收據設定為1.個人
'   'Modify By Sindy 2025/6/23 and a0k05='1' 改為 and a0k05='2' 2=可扣繳
'   cnnConnection.Execute "delete from ACCTMP44q0 where T05='" & Me.Name & "' and T14='" & strUserNum & "'", intI
'   strExc(0) = "SELECT substr(a0k01,1,1) as ID,a0k01,A1v03,A1v09,A0k04,a0k03 as Cuno,sum(A1v06),'" & Me.Name & "','" & strUserNum & "'" & _
'               " From acc0k0, acc1v0" & _
'               " where a0k01=a1v02(+) and a0k05='2' and a1v09=" & txtYear & _
'               " and length(A0k04)>=4" & _
'               " group by a0k01,A1v03,A1v09,A0k04,a0k03" & _
'               " Union" & _
'               " SELECT substr(a1k01,1,1) as ID,a1k01 as a0k01,A1v03,A1v09,A1k35 as A0k04,a1k28 as Cuno,sum(A1v06),'" & Me.Name & "','" & strUserNum & "'" & _
'               " From acc1k0, acc1v0" & _
'               " where a1k01=a1v02(+) and a1v09=" & txtYear & _
'               " and length(A1k35)>=4" & _
'               " group by a1k01,A1v03,A1v09,A1k35,a1k28"
'   cnnConnection.Execute "insert into ACCTMP44q0(T06,T01,T07,T03,T15,T02,T04,T05,T14) " & strExc(0), intI
   'Modify By Sindy 2025/6/26 使用共用函數抓好的資料
   lngPageNo = 0
'   strExc(0) = "SELECT T06,T01,T07,T03,T15,T02,T04,cu158" & _
'               " From ACCTMP44q0,customer" & _
'               " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
'               " and substr(T02,1,8)=cu01(+) and substr(T02,9,1)=cu02(+)" & _
'               " order by T06,T15,T01"
   
'   '************************************************************************************
'   '暫存檔
'   'T01 : 收據號碼
'   'T02 : 客戶編號 ==> 改用"T29"收據抬據抓出來的客戶編號
'   'T03 : 扣繳年度
'   'T04 : 已扣繳金額
'   'T05 : Form Name
'   'T06 : X.Acc1K0 或 E.Acc0K0
'   'T07 : 公司別
'   'T14 : UserID
'   'T15 : 收據抬頭
'   'T25 : 是否為境外公司
'   '************************************************************************************
   strExc(0) = "SELECT substr(T01,1,1),T01,T07,T03,T15,nvl(t29,t02) as T02,T04,T25" & _
               " From ACCTMP44q0" & _
               " where T05='" & Me.Name & "' and T14='" & strUserNum & "'" & _
               " order by substr(T01,1,1),T15,T01"
   '2025/6/26 END
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoquery
         .MoveFirst
         Set xlsAnnuity = New Excel.Application
         Call SetExcelWorksheets
         PrintHead_Excel intCounter '頁首
         Do While Not .EOF
'            '第2頁切頁有誤 +  And intCounter <> 48 判斷
'            If (lngPageNo = 1 And intCounter Mod 32 = 0) Or _
'               (lngPageNo <> 1 And intCounter Mod 32 = 0 And intCounter <> 32) Then
'               '換頁
'               intCounter = intCounter + 1
'               wksAnnuity.Range("A" & intCounter).Select
'               wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
'               PrintHead_Excel intCounter '頁首
'            End If
            '明細資料
            PrintData_Excel adoquery, intCounter
            .MoveNext
         Loop
      End With
   Else
      Process = False
      Screen.MousePointer = vbDefault
      'MsgBox "扣繳收據抬頭檢查清單；無資料，可供列印！"
      adoquery.Close
      Set adoquery = Nothing
      Exit Function
   End If
   
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   'MsgBox "資料產生完畢！"
   Process = True
   
   Screen.MousePointer = vbDefault
   
   Set adoquery = Nothing
   Exit Function

ErrHnd:
   Screen.MousePointer = vbDefault
   Set adoquery = Nothing
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetExcelWorksheets()
   xlsAnnuity.Visible = False 'True
   xlsAnnuity.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 10
   wksAnnuity.Columns("B:B").ColumnWidth = 6
   wksAnnuity.Columns("C:C").ColumnWidth = 10
   wksAnnuity.Columns("D:D").ColumnWidth = 50
   wksAnnuity.Columns("E:E").ColumnWidth = 10
   wksAnnuity.Columns("F:F").ColumnWidth = 10
   wksAnnuity.Columns("G:G").ColumnWidth = 10
   
   wksAnnuity.Range("A:A").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("B:B").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("C:C").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("D:D").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("E:E").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("F:F").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("G:G").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   
   intCounter = 1
End Sub

'表頭
Private Sub PrintHead_Excel(ByRef iRow As Integer)
Dim i As Integer, strTemp As String

   lngPageNo = lngPageNo + 1
   With wksAnnuity
      .Range("E" & iRow).Value = "扣繳收據抬頭檢查清單"
      '選取,儲存格合併,置中,粗體字
      strTemp = "A" & iRow & ":F" & iRow
      .Range(strTemp).Select
      With .Application.Selection
          .HorizontalAlignment = xlGeneral
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With .Application.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      .Application.Selection.Font.Bold = True

      iRow = iRow + 1
      .Range("A" & iRow).Value = "列印人：" & strUserName
      .Range("C" & iRow).Value = "扣繳年度：" & txtYear
      .Range("E" & iRow).Value = "列印日期："
      .Range("F" & iRow).Value = Format(strSrvDate(2), "###/##/##")
      iRow = iRow + 1
'      .Range("E" & iRow).Value = "頁數："
'      .Range("F" & iRow).Value = lngPageNo
'      strTemp = "D" & iRow - 1 & ":D" & iRow
'      .Range(strTemp).Select
'      With .Application.Selection
'         .HorizontalAlignment = xlCenter '置中
'      End With
      strTemp = "E" & iRow - 1 & ":E" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlRight '靠右
      End With
      strTemp = "F" & iRow & ":F" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlLeft '靠左
      End With
      
      iRow = iRow + 1
      .Range("A" & iRow).Value = "收據編號"
      .Range("B" & iRow).Value = "公司別"
      .Range("C" & iRow).Value = "扣繳年度"
      .Range("D" & iRow).Value = "收據抬頭"
      .Range("E" & iRow).Value = "客戶編號"
      .Range("F" & iRow).Value = "是否境外"
      .Range("G" & iRow).Value = "已扣繳金額"
      strTemp = "A" & iRow & ":G" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
'      With .Application.Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeTop)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
'      With .Application.Selection.Borders(xlEdgeRight)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlInsideVertical)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
   End With
End Sub

Private Sub PrintData_Excel(p_Rst As ADODB.Recordset, ByRef iRow As Integer)
Dim strTemp As String
   
   iRow = iRow + 1
   With wksAnnuity
      .Range("A" & iRow).Value = "" & p_Rst.Fields("T01")
      .Range("B" & iRow).Value = "" & p_Rst.Fields("T07")
      .Range("C" & iRow).Value = "" & p_Rst.Fields("T03")
      .Range("D" & iRow).Value = "" & p_Rst.Fields("T15")
      .Range("E" & iRow).Value = "" & p_Rst.Fields("T02")
      .Range("F" & iRow).Value = "" & p_Rst.Fields("T25")
      .Range("G" & iRow).Value = "" & p_Rst.Fields("T04")
'      .Range("H" & iRow).Value = "" & p_Rst.Fields("T26") '會計備註
'      .Range("A" & iRow & ":G" & iRow).Select
'      .Application.Selection.VerticalAlignment = xlTop '靠上
'      .Range("H" & iRow).Select
'      .Application.Selection.WrapText = True '自動換行

'      strTemp = "A" & iRow & ":I" & iRow
'      .Range(strTemp).Select
'      With .Application.Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeTop)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeBottom)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeRight)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlInsideVertical)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
   End With
End Sub

Private Sub Form_Activate()
   strFormName = Name
End Sub

Private Sub Form_Load()
   '表單初始化
   PUB_InitForm Me, 5640, 4450, strBackPicPath4
   '扣繳年月預設為今年
   txtYear = Left(strSrvDate(2), 3)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc44x0 = Nothing
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If txtYear = "" Then
      MsgBox "請輸入扣繳年度！", vbExclamation
      FormCheck = False
      txtYear.SetFocus
      Exit Function
   End If
   
   If Check1.Value = 0 And Check2.Value = 0 And Check3.Value = 0 Then
      MsgBox "請至少勾選一項檢核清單！", vbExclamation
      FormCheck = False
      Exit Function
   End If
   
   FormCheck = True
End Function

Private Function Process2() As Boolean
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   
'   '有扣繳但沒財務信箱cu115,也沒有代表信箱cu20,也沒有會計師信箱A4905
'   'Modify By Sindy 2025/6/17 +, FAGENT
'   '                          +AND not exists...
'   strExc(0) = "SELECT DISTINCT ST15,A0902 業務區,ST01,ST02 智權人員,NVL(A0K03,A1K28) 客戶編號,nvl(a0k04,a1k35) 收據抬頭," & _
'               "decode(DECODE(CU16||CU17,NULL,A4204,CU16||decode(cu17,NULL,NULL,';'||CU17)),null,fa12||decode(fa13,NULL,NULL,';'||fa13)) 電話" & _
'               " From acc1v0, acc1k0, acc0k0, customer, acc420, STAFF, ACC090, ACC490, FAGENT" & _
'               " Where a1v09=" & txtYear & " And NVL(a1v06, 0) > 0" & _
'               " AND a1v02=a0k01(+) AND a1v02=a1k01(+)" & _
'               " AND nvl(a0k04,a1k35)=cu04(+) AND nvl(a0k04,a1k35)=a4201(+)" & _
'               " AND NVL(CU13,A4206)=ST01(+) AND ST15=A0901(+)" & _
'               " AND decode(cu01,NULL,a4201,cu01||'0')=a4901(+)" & _
'               " AND a4218||CU20||cu115||a4905||fa16||fa79 IS NULL" & _
'               " AND substr(NVL(A0K03,A1K28),1,8)=fa01(+) AND substr(NVL(A0K03,A1K28),9,1)=fa02(+)" & _
'               " AND not exists(select C2.CU01||C2.CU02 from customer C2, ACC490" & _
'               " where C2.CU01||C2.CU02 in(select min(C1.CU01||C1.CU02) from customer C1 where a0k04=C1.cu04 AND C1.cu02||''='0')" & _
'               " AND C2.CU01||C2.CU02=a4901(+) AND C2.CU20||C2.cu115||a4905 is not null)" & _
'               " order by ST15,ST01,NVL(A0K03,A1K28)"
   'Modify By Sindy 2025/6/27 使用共用函數抓好的資料
   lngPageNo = 0
   strExc(0) = "SELECT distinct ST15,A0902 業務區,ST01,ST02 智權人員,nvl(t29,t02) 客戶編號,T15 收據抬頭,T17 電話,decode(T29,'T','Y',null) T29" & _
               " From ACCTMP44q0, STAFF, ACC090" & _
               " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T04>0" & _
               " AND T23=ST01(+) AND ST15=A0901(+)" & _
               " AND (T06||T16||T20 IS NULL or T06||T16||T20='X')" & _
               " order by ST15,ST01,客戶編號"
   '2025/6/27 END
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoquery
         .MoveFirst
         Set xlsAnnuity = New Excel.Application
         Call SetExcelWorksheets2
         PrintHead_Excel2 intCounter '頁首
         Do While Not .EOF
            '明細資料
            PrintData_Excel2 adoquery, intCounter
            .MoveNext
         Loop
      End With
   Else
      Process2 = False
      Screen.MousePointer = vbDefault
      'MsgBox "代表、財務、會計師信箱均無建立；無資料，可供列印！"
      adoquery.Close
      Set adoquery = Nothing
      Exit Function
   End If
   
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   'MsgBox "資料產生完畢！"
   Process2 = True
   
   Screen.MousePointer = vbDefault
   
   Set adoquery = Nothing
   Exit Function

ErrHnd:
   Screen.MousePointer = vbDefault
   Set adoquery = Nothing
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetExcelWorksheets2()
   xlsAnnuity.Visible = False 'True
   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 10
   wksAnnuity.Columns("B:B").ColumnWidth = 10
   wksAnnuity.Columns("C:C").ColumnWidth = 10
   wksAnnuity.Columns("D:D").ColumnWidth = 10
   wksAnnuity.Columns("E:E").ColumnWidth = 10
   wksAnnuity.Columns("F:F").ColumnWidth = 40
   wksAnnuity.Columns("G:G").ColumnWidth = 10
   wksAnnuity.Columns("H:H").ColumnWidth = 12
   
   wksAnnuity.Range("A:A").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("B:B").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("C:C").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("D:D").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("E:E").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("F:F").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("G:G").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("H:H").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   
   intCounter = 1
End Sub

'表頭
Private Sub PrintHead_Excel2(ByRef iRow As Integer)
Dim i As Integer, strTemp As String

   lngPageNo = lngPageNo + 1
   With wksAnnuity
      .Range("E" & iRow).Value = "代表、財務、會計師信箱均無建立"
      '選取,儲存格合併,置中,粗體字
      strTemp = "A" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
          .HorizontalAlignment = xlGeneral
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With .Application.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      .Application.Selection.Font.Bold = True

      iRow = iRow + 1
      .Range("A" & iRow).Value = "列印人：" & strUserName
      .Range("C" & iRow).Value = "扣繳年度：" & txtYear
      .Range("F" & iRow).Value = "列印日期："
      .Range("G" & iRow).Value = Format(strSrvDate(2), "###/##/##")
      iRow = iRow + 1
'      .Range("F" & iRow).Value = "頁數："
'      .Range("G" & iRow).Value = lngPageNo
'      strTemp = "D" & iRow - 1 & ":D" & iRow
'      .Range(strTemp).Select
'      With .Application.Selection
'         .HorizontalAlignment = xlCenter '置中
'      End With
      strTemp = "F" & iRow - 1 & ":F" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlRight '靠右
      End With
      strTemp = "G" & iRow & ":G" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlLeft '靠左
      End With
      strTemp = "H" & iRow - 1 & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlRight '靠右
      End With
      
      iRow = iRow + 1
      .Range("A" & iRow).Value = "業務區代碼"
      .Range("B" & iRow).Value = "業務區"
      .Range("C" & iRow).Value = "智權人員編號"
      .Range("D" & iRow).Value = "智權人員"
      .Range("E" & iRow).Value = "客戶編號"
      .Range("F" & iRow).Value = "收據抬頭"
      .Range("G" & iRow).Value = "電話"
      .Range("H" & iRow).Value = "特殊收據抬頭" 'Add By Sindy 2025/6/30
      strTemp = "A" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
'      With .Application.Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeTop)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
'      With .Application.Selection.Borders(xlEdgeRight)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlInsideVertical)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
   End With
End Sub

Private Sub PrintData_Excel2(p_Rst As ADODB.Recordset, ByRef iRow As Integer)
Dim strTemp As String
   
   iRow = iRow + 1
   With wksAnnuity
      .Range("A" & iRow).Value = "" & p_Rst.Fields("ST15")
      .Range("B" & iRow).Value = "" & p_Rst.Fields("業務區")
      .Range("C" & iRow).Value = "" & p_Rst.Fields("ST01")
      .Range("D" & iRow).Value = "" & p_Rst.Fields("智權人員")
      .Range("E" & iRow).Value = "" & p_Rst.Fields("客戶編號")
      .Range("F" & iRow).Value = "" & p_Rst.Fields("收據抬頭")
      .Range("G" & iRow).Value = "" & p_Rst.Fields("電話")
      .Range("H" & iRow).Value = "" & p_Rst.Fields("T29") 'Add By Sindy 2025/6/30
   End With
End Sub

Private Function Process3() As Boolean
   
On Error GoTo ErrHnd
   
   Screen.MousePointer = vbHourglass
   
'   '有扣繳有代表信箱但沒有財務信箱,也沒有會計師信箱A4905
'   strExc(0) = "SELECT DISTINCT ST15,A0902 業務區,ST01,ST02 智權人員,NVL(A0K03,A1K28) 客戶編號,nvl(a0k04,a1k35) 收據抬頭," & _
'               "CU16||decode(cu17,null,null,';'||CU17) 電話" & _
'               " From acc1v0, acc1k0, acc0k0, customer, STAFF, ACC090, ACC490" & _
'               " Where a1v09=" & txtYear & " And nvl(a1v06, 0) > 0" & _
'               " AND a1v02=a0k01(+) AND a1v02=a1k01(+)" & _
'               " AND nvl(a0k04,a1k35)=cu04(+) AND CU01 IS NOT NULL" & _
'               " AND CU20 IS NOT NULL AND CU13=ST01(+) AND ST15=A0901(+)" & _
'               " AND cu01||'0'=a4901(+) AND CU115||a4905 IS NULL" & _
'               " order by ST15,ST01,NVL(A0K03,A1K28)"
   'Modify By Sindy 2025/6/27 使用共用函數抓好的資料
   lngPageNo = 0
   strExc(0) = "SELECT distinct ST15,A0902 業務區,ST01,ST02 智權人員,nvl(t29,t02) 客戶編號,T15 收據抬頭,T17 電話,decode(T29,'T','Y',null) T29" & _
               " From ACCTMP44q0, STAFF, ACC090" & _
               " where T05='" & Me.Name & "' and T14='" & strUserNum & "' and T04>0" & _
               " AND T23=ST01(+) AND ST15=A0901(+)" & _
               " AND T06 IS NOT NULL AND T06<>'X' AND T16||T20 IS NULL" & _
               " order by ST15,ST01,客戶編號"
   '2025/6/27 END
   intI = 1
   Set adoquery = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With adoquery
         .MoveFirst
         Set xlsAnnuity = New Excel.Application
         Call SetExcelWorksheets3
         PrintHead_Excel3 intCounter '頁首
         Do While Not .EOF
            '明細資料
            PrintData_Excel3 adoquery, intCounter
            .MoveNext
         Loop
      End With
   Else
      Process3 = False
      Screen.MousePointer = vbDefault
      'MsgBox "有代表信箱，但無財務、會計師信箱；無資料，可供列印！"
      adoquery.Close
      Set adoquery = Nothing
      Exit Function
   End If
   
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing
   'MsgBox "資料產生完畢！"
   Process3 = True
   
   Screen.MousePointer = vbDefault
   
   Set adoquery = Nothing
   Exit Function

ErrHnd:
   Screen.MousePointer = vbDefault
   Set adoquery = Nothing
   xlsAnnuity.Visible = True
   xlsAnnuity.WindowState = wdWindowStateMaximize
   Set xlsAnnuity = Nothing
   Set wksAnnuity = Nothing

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SetExcelWorksheets3()
   xlsAnnuity.Visible = False 'True
   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.RightMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.TopMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.BottomMargin = 42.51 'Application.InchesToPoints(0.590551181102362)
   wksAnnuity.PageSetup.HeaderMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   wksAnnuity.PageSetup.FooterMargin = 28.34 'Application.InchesToPoints(0.393700787401575)
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 10
   wksAnnuity.Columns("B:B").ColumnWidth = 10
   wksAnnuity.Columns("C:C").ColumnWidth = 10
   wksAnnuity.Columns("D:D").ColumnWidth = 10
   wksAnnuity.Columns("E:E").ColumnWidth = 10
   wksAnnuity.Columns("F:F").ColumnWidth = 40
   wksAnnuity.Columns("G:G").ColumnWidth = 10
   wksAnnuity.Columns("H:H").ColumnWidth = 12
   
   wksAnnuity.Range("A:A").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("B:B").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("C:C").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("D:D").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("E:E").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("F:F").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("G:G").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   wksAnnuity.Range("H:H").Select
   wksAnnuity.Application.Selection.NumberFormatLocal = "@" '文字
   
   intCounter = 1
End Sub

'表頭
Private Sub PrintHead_Excel3(ByRef iRow As Integer)
Dim i As Integer, strTemp As String

   lngPageNo = lngPageNo + 1
   With wksAnnuity
      .Range("E" & iRow).Value = "有代表信箱，但無財務、會計師信箱"
      '選取,儲存格合併,置中,粗體字
      strTemp = "A" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
          .HorizontalAlignment = xlGeneral
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With .Application.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      .Application.Selection.Font.Bold = True

      iRow = iRow + 1
      .Range("A" & iRow).Value = "列印人：" & strUserName
      .Range("C" & iRow).Value = "扣繳年度：" & txtYear
      .Range("F" & iRow).Value = "列印日期："
      .Range("G" & iRow).Value = Format(strSrvDate(2), "###/##/##")
      iRow = iRow + 1
'      .Range("F" & iRow).Value = "頁數："
'      .Range("G" & iRow).Value = lngPageNo
'      strTemp = "D" & iRow - 1 & ":D" & iRow
'      .Range(strTemp).Select
'      With .Application.Selection
'         .HorizontalAlignment = xlCenter '置中
'      End With
      strTemp = "F" & iRow - 1 & ":F" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlRight '靠右
      End With
      strTemp = "G" & iRow & ":G" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlLeft '靠左
      End With
      strTemp = "H" & iRow - 1 & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlRight '靠右
      End With
      
      iRow = iRow + 1
      .Range("A" & iRow).Value = "業務區代碼"
      .Range("B" & iRow).Value = "業務區"
      .Range("C" & iRow).Value = "智權人員編號"
      .Range("D" & iRow).Value = "智權人員"
      .Range("E" & iRow).Value = "客戶編號"
      .Range("F" & iRow).Value = "收據抬頭"
      .Range("G" & iRow).Value = "電話"
      .Range("H" & iRow).Value = "特殊收據抬頭" 'Add By Sindy 2025/6/30
      strTemp = "A" & iRow & ":H" & iRow
      .Range(strTemp).Select
      With .Application.Selection
         .HorizontalAlignment = xlCenter '置中
      End With
'      With .Application.Selection.Borders(xlEdgeLeft)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlEdgeTop)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
      With .Application.Selection.Borders(xlEdgeBottom)
         .LineStyle = xlContinuous
         .Weight = xlThin
         .ColorIndex = xlAutomatic
      End With
'      With .Application.Selection.Borders(xlEdgeRight)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
'      With .Application.Selection.Borders(xlInsideVertical)
'         .LineStyle = xlContinuous
'         .Weight = xlThin
'         .ColorIndex = xlAutomatic
'      End With
   End With
End Sub

Private Sub PrintData_Excel3(p_Rst As ADODB.Recordset, ByRef iRow As Integer)
Dim strTemp As String
   
   iRow = iRow + 1
   With wksAnnuity
      .Range("A" & iRow).Value = "" & p_Rst.Fields("ST15")
      .Range("B" & iRow).Value = "" & p_Rst.Fields("業務區")
      .Range("C" & iRow).Value = "" & p_Rst.Fields("ST01")
      .Range("D" & iRow).Value = "" & p_Rst.Fields("智權人員")
      .Range("E" & iRow).Value = "" & p_Rst.Fields("客戶編號")
      .Range("F" & iRow).Value = "" & p_Rst.Fields("收據抬頭")
      .Range("G" & iRow).Value = "" & p_Rst.Fields("電話")
      .Range("H" & iRow).Value = "" & p_Rst.Fields("T29") 'Add By Sindy 2025/6/30
   End With
End Sub
