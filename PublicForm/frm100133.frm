VERSION 5.00
Begin VB.Form frm100133 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報產業分類案件市佔分析"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5355
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   5355
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   435
      Left            =   1980
      TabIndex        =   16
      Top             =   1680
      Width           =   2205
      Begin VB.OptionButton Option2 
         Caption         =   "產業別"
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   4
         Top             =   60
         Value           =   -1  'True
         Width           =   885
      End
      Begin VB.OptionButton Option2 
         Caption         =   "案件屬性"
         Height          =   345
         Index           =   1
         Left            =   960
         TabIndex        =   5
         Top             =   60
         Width           =   1185
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      Height          =   435
      Left            =   1980
      TabIndex        =   14
      Top             =   1200
      Width           =   2205
      Begin VB.OptionButton Option1 
         Caption         =   "公開公報"
         Height          =   345
         Index           =   1
         Left            =   960
         TabIndex        =   3
         Top             =   60
         Width           =   1185
      End
      Begin VB.OptionButton Option1 
         Caption         =   "公報"
         Height          =   345
         Index           =   0
         Left            =   30
         TabIndex        =   2
         Top             =   60
         Value           =   -1  'True
         Width           =   885
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   3030
      MaxLength       =   5
      TabIndex        =   1
      Top             =   870
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   210
      TabIndex        =   8
      Top             =   2670
      Visible         =   0   'False
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   9
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   10
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1980
      MaxLength       =   5
      TabIndex        =   0
      Top             =   870
      Width           =   825
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   375
      Index           =   1
      Left            =   4245
      TabIndex        =   7
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "產生Excel檔(&E)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   6
      Top             =   90
      Width           =   1455
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "統計分類："
      Height          =   180
      Left            =   1050
      TabIndex        =   17
      Top             =   1830
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "公報種類："
      Height          =   180
      Left            =   1050
      TabIndex        =   15
      Top             =   1350
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "註：產業分類資料開始於101年01月。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   930
      TabIndex        =   13
      Top             =   2460
      Width           =   2970
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "~"
      Height          =   180
      Left            =   2850
      TabIndex        =   12
      Top             =   900
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "起迄年月："
      Height          =   180
      Left            =   1050
      TabIndex        =   11
      Top             =   900
      Width           =   900
   End
End
Attribute VB_Name = "frm100133"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/19 Form2.0已檢查 (無需修改的物件)
'Create by Sindy 2013/8/28
Option Explicit

Dim strStarYM As String, strEndYM As String
Dim strCountYear As String, strCountMon As String, strChkData As String
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim lngCounter As Long, i As Integer, j As Integer
Dim strFileName As String, strKind As String
Dim dblVal(1 To 100) As Double 'Modify By Sindy 2016/2/26
Dim dblValTot(1 To 100) As Double 'Modify By Sindy 2016/2/26
Dim intRunCnt As Integer, intRunItem As Integer 'Add By Sindy 2016/2/26
Dim strConSql As String 'Add By Sindy 2016/2/26


Private Sub cmdOK_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Or txt1(1) = "" Then
            MsgBox "起迄年月不可空白！", vbInformation, "操作錯誤！"
            If txt1(0) = "" Then txt1(0).SetFocus
            If txt1(1) = "" Then txt1(1).SetFocus
            Exit Sub
         End If
         
         strStarYM = Val(txt1(0)) + 191100
         strEndYM = Val(txt1(1)) + 191100
         If strStarYM > strEndYM Then
            MsgBox "起始年月不可大於終止年月！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
         End If
         '逐月檢查有無資料
         strCountYear = Left(strStarYM, 4)
         strCountMon = Right(strStarYM, 2)
         strChkData = ""
         If Option1(0).Value = True Then '公報
            Do While Val(strCountYear & strCountMon) <= Val(strEndYM)
               'Modify By Sindy 2016/3/2
               If Option2(1).Value = True Then '案件屬性
                  strExc(0) = "select count(*) from tpbulletin Where TPB03>=" & Val(strCountYear & strCountMon) & "01 and TPB03<=" & Val(strCountYear & strCountMon) & "31 and TPB13 is not null"
               Else
               '2016/3/2 END
                  strExc(0) = "select count(*) from tpbulletin Where TPB03>=" & Val(strCountYear & strCountMon) & "01 and TPB03<=" & Val(strCountYear & strCountMon) & "31 and TPB12 is not null"
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If IsNull(RsTemp.Fields(0)) = True Then
                     strChkData = strChkData & "、" & strCountYear & strCountMon
                  ElseIf RsTemp.Fields(0) = 0 Then
                     strChkData = strChkData & "、" & strCountYear & strCountMon
                  End If
               End If
               strCountMon = Right("0" & CStr(Val(strCountMon) + 1), 2)
               If Val(strCountMon) > 12 Then
                  strCountYear = Val(strCountYear) + 1
                  strCountMon = "01"
               End If
            Loop
         Else '公開公報
            Do While Val(strCountYear & strCountMon) <= Val(strEndYM)
               'Modify By Sindy 2016/3/2
               If Option2(1).Value = True Then '案件屬性
                  strExc(0) = "select count(*) from tpgazette Where TPG03>=" & Val(strCountYear & strCountMon) & "01 and TPG03<=" & Val(strCountYear & strCountMon) & "31 and TPG18 is not null"
               Else
               '2016/3/2 END
                  strExc(0) = "select count(*) from tpgazette Where TPG03>=" & Val(strCountYear & strCountMon) & "01 and TPG03<=" & Val(strCountYear & strCountMon) & "31 and TPG17 is not null"
               End If
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If IsNull(RsTemp.Fields(0)) = True Then
                     strChkData = strChkData & "、" & strCountYear & strCountMon
                  ElseIf RsTemp.Fields(0) = 0 Then
                     strChkData = strChkData & "、" & strCountYear & strCountMon
                  End If
               End If
               strCountMon = Right("0" & CStr(Val(strCountMon) + 1), 2)
               If Val(strCountMon) > 12 Then
                  strCountYear = Val(strCountYear) + 1
                  strCountMon = "01"
               End If
            Loop
         End If
         If strChkData <> "" Then
            strChkData = Mid(strChkData, 2)
            If strStarYM = strEndYM Then
               MsgBox "查無資料!!!", vbExclamation + vbOKOnly
               txt1(0).SetFocus
               Exit Sub
            Else
               If MsgBox(strChkData & "查無資料，確定還要進行統計嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Exit Sub
               End If
            End If
         End If
         
         Screen.MousePointer = vbHourglass
         
         'Add By Sindy 2016/2/26
         intRunCnt = 16 '幾家事務所(含全國) * 2
         If Option2(0) Then
            intRunItem = 36 '+ 其他
         Else
            intRunItem = 4 '+ 其他
         End If
         '2016/2/26 END
         
         If Option1(0).Value = True Then
            Call ExcelSave '公報
         Else
            Call ExcelSave_2 '公開公報
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

'*************************************************
'  轉成Excel檔案 (公報)
'
'*************************************************
Private Sub ExcelSave()
On Error GoTo flgErr
   
   'Modify By Sindy 2016/2/26 產業別
   If Option2(0).Value = True Then
      If strStarYM <> strEndYM Then
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月至" & Left(strEndYM, 4) & "年" & Right(strEndYM, 2) & "月（公報）產業分類案件市佔分析.xls"
      Else
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月（公報）產業分類案件市佔分析.xls"
      End If
   Else '案件屬性
      If strStarYM <> strEndYM Then
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月至" & Left(strEndYM, 4) & "年" & Right(strEndYM, 2) & "月（公報）案件屬性案件市佔分析.xls"
      Else
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月（公報）案件屬性案件市佔分析.xls"
      End If
   End If
   '2016/2/26 END
   
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   xlsSalesPoint.SheetsInNewWorkbook = 3 'Added by Lydia 2019/03/13 預設工作表數量
   xlsSalesPoint.Workbooks.add
   
   Call ExcelSave_Total
   Call ExcelSave_FCP
   Call ExcelSave_CCP
   
   'Modify By Sindy 2018/5/9
   'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   '2018/5/9 END
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   MsgBox "檔案已產生！電子檔位置：" & strFileName
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub ExcelSave_Total()
On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
   wksaccrpt114.Name = "Total"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
'   wksaccrpt114.Range("a1:g1").Select
'   With wksaccrpt114.Range("a1:g1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   
   'Modify By Sindy 2016/2/26
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   '2016/2/26 END
    
   For i = 1 To intRunItem '36 分類項目
      'Modify By Sindy 2016/2/26
      'If i = 36 Then i = 99 '其他
      If Option2(0).Value = True Then '產業別
         If i = intRunItem Then i = 99 '其他
         strKind = GetItemNm(i)
         strConSql = " and TPB12 is not null and TPB12='" & Format(i, "00") & "'"
      Else '案件屬性
         If i = intRunItem Then i = 9 '其他
         strKind = GetItemNm2(i)
         strConSql = " and TPB13 is not null and TPB13='" & Format(i, "0") & "'"
      End If
      '2016/2/26 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M'"
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1'"
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1'"
      'Add By Sindy 2016/2/26
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='將群'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='將群'"
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='冠群國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='冠群國際'"
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='連邦國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='連邦國際'"
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='聖島國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='聖島國際'"
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='理律法律'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='理律法律'"
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='台灣國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='台灣國際'"
      '2016/2/26 END
      strSql = strSql & ")"
      
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      '2016/2/26 END
      
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
'         dblVal1 = adoRecordset.Fields(0)
'         dblVal2 = adoRecordset.Fields(1)
'         dblVal3 = adoRecordset.Fields(2)
'         dblVal4 = adoRecordset.Fields(3)
         'Modify By Sindy 2016/2/26
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/2/26 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      'Modify By Sindy 2016/2/26
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j) '3
         wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1) '4
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
         If dblVal(2) = 0 Then
            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
         End If
         j = j + 1
         lngCounter = lngCounter + 1
      Next j
      '2016/2/26 END
      '合計
'      dblValTot_1 = dblValTot_1 + dblVal1
'      dblValTot_2 = dblValTot_2 + dblVal2
'      dblValTot_3 = dblValTot_3 + dblVal3
'      dblValTot_4 = dblValTot_4 + dblVal4
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2016/2/26 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   'Modify By Sindy 2016/2/26
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j) '3
      wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1) '4
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
      If dblValTot(2) = 0 Then
         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
      End If
      j = j + 1
      lngCounter = lngCounter + 1
   Next j
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub ExcelSave_FCP()
On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(2)
   wksaccrpt114.Name = "FCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
'   wksaccrpt114.Range("a1:g1").Select
'   With wksaccrpt114.Range("a1:g1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   
   'Modify By Sindy 2016/2/26
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   '2016/2/26 END
   
   For i = 1 To intRunItem '36 分類項目
      'Modify By Sindy 2016/2/26
      'If i = 36 Then i = 99 '其他
      If Option2(0).Value = True Then '產業別
         If i = intRunItem Then i = 99 '其他
         strKind = GetItemNm(i)
         strConSql = " and TPB12 is not null and TPB12='" & Format(i, "00") & "'"
      Else '案件屬性
         If i = intRunItem Then i = 9 '其他
         strKind = GetItemNm2(i)
         strConSql = " and TPB13 is not null and TPB13='" & Format(i, "0") & "'"
      End If
      '2016/2/26 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M'"
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1'"
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1'"
      'Add By Sindy 2016/2/26
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='將群'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='將群'"
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='冠群國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='冠群國際'"
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='連邦國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='連邦國際'"
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='聖島國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='聖島國際'"
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='理律法律'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='理律法律'"
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='台灣國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and TPB06>'010' " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='台灣國際'"
      '2016/2/26 END
      strSql = strSql & ")"
      
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      '2016/2/26 END
      
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
'         dblVal1 = adoRecordset.Fields(0)
'         dblVal2 = adoRecordset.Fields(1)
'         dblVal3 = adoRecordset.Fields(2)
'         dblVal4 = adoRecordset.Fields(3)
         'Modify By Sindy 2016/2/26
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/2/26 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      'Modify By Sindy 2016/2/26
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j) '3
         wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1) '4
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
         If dblVal(2) = 0 Then
            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
         End If
         j = j + 1
         lngCounter = lngCounter + 1
      Next j
      '2016/2/26 END
      '合計
'      dblValTot_1 = dblValTot_1 + dblVal1
'      dblValTot_2 = dblValTot_2 + dblVal2
'      dblValTot_3 = dblValTot_3 + dblVal3
'      dblValTot_4 = dblValTot_4 + dblVal4
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2016/2/26 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   'Modify By Sindy 2016/2/26
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
      wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
      If dblValTot(2) = 0 Then
         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
      End If
      j = j + 1
      lngCounter = lngCounter + 1
   Next j
   '2016/2/26 END
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub ExcelSave_CCP()
On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(3)
   wksaccrpt114.Name = "CCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
'   wksaccrpt114.Range("a1:g1").Select
'   With wksaccrpt114.Range("a1:g1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   
   'Modify By Sindy 2016/2/26
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   '2016/2/26 END
   
   For i = 1 To intRunItem '36
      'Modify By Sindy 2016/2/26
      'If i = 36 Then i = 99 '其他
      If Option2(0).Value = True Then '產業別
         If i = intRunItem Then i = 99 '其他
         strKind = GetItemNm(i)
         strConSql = " and TPB12 is not null and TPB12='" & Format(i, "00") & "'"
      Else '案件屬性
         If i = intRunItem Then i = 9 '其他
         strKind = GetItemNm2(i)
         strConSql = " and TPB13 is not null and TPB13='" & Format(i, "0") & "'"
      End If
      '2016/2/26 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M'"
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1'"
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1'"
      'Add By Sindy 2016/2/26
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='將群'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='將群'"
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='冠群國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='冠群國際'"
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='連邦國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='連邦國際'"
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='聖島國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='聖島國際'"
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='理律法律'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='理律法律'"
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='I' and TPB08='台灣國際'"
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and (TPB06<='010' or TPB06 is null) " & strConSql & " and substr(TPB02,1,1)='M' and TPB08='台灣國際'"
      '2016/2/26 END
      strSql = strSql & ")"
      
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      '2016/2/26 END
      
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
'         dblVal1 = adoRecordset.Fields(0)
'         dblVal2 = adoRecordset.Fields(1)
'         dblVal3 = adoRecordset.Fields(2)
'         dblVal4 = adoRecordset.Fields(3)
         'Modify By Sindy 2016/2/26
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/2/26 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      'Modify By Sindy 2016/2/26
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j) '3
         wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1) '4
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
         If dblVal(2) = 0 Then
            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
         End If
         j = j + 1
         lngCounter = lngCounter + 1
      Next j
      '2016/2/26 END
      '合計
'      dblValTot_1 = dblValTot_1 + dblVal1
'      dblValTot_2 = dblValTot_2 + dblVal2
'      dblValTot_3 = dblValTot_3 + dblVal3
'      dblValTot_4 = dblValTot_4 + dblVal4
      'Modify By Sindy '2016/2/26
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      ''2016/2/26 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   'Modify By Sindy 2016/2/26
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j) '3
      wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1) '4
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
      If dblValTot(2) = 0 Then
         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
      End If
      j = j + 1
      lngCounter = lngCounter + 1
   Next j
   '2016/2/26 END
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'產業別
Private Function GetItemNm(intItem As Integer) As String
   GetItemNm = ""
   Select Case intItem
      Case 1
         GetItemNm = "農林漁牧"
      Case 2
         GetItemNm = "食品及煙草"
      Case 3
         GetItemNm = "日常用品"
      Case 4
         GetItemNm = "保健及娛樂"
      Case 5
         GetItemNm = "生物技術"
      Case 6
         GetItemNm = "醫藥品"
      Case 7
         GetItemNm = "分離及混合"
      Case 8
         GetItemNm = "成型"
      Case 9
         GetItemNm = "印刷"
      Case 10
         GetItemNm = "運輸"
      Case 11
         GetItemNm = "微型結構技術、超微技術"
      Case 12
         GetItemNm = "無機化學、廢水處理"
      Case 13
         GetItemNm = "有機化學"
      Case 14
         GetItemNm = "高分子"
      Case 15
         GetItemNm = "染料、石油、動植物油"
      Case 16
         GetItemNm = "糖，皮革"
      Case 17
         GetItemNm = "冶金、金屬表面處理、電鍍"
      Case 18
         GetItemNm = "紡織及不屬別類之柔性材料"
      Case 19
         GetItemNm = "造紙及紙製品加工"
      Case 20
         GetItemNm = "土木建築"
      Case 21
         GetItemNm = "採礦"
      Case 22
         GetItemNm = "引擎及泵"
      Case 23
         GetItemNm = "一般機械工程"
      Case 24
         GetItemNm = "照明；加熱"
      Case 25
         GetItemNm = "武器；爆破"
      Case 26
         GetItemNm = "儀器1（光學）"
      Case 27
         GetItemNm = "儀器2（量測）"
      Case 28
         GetItemNm = "儀器３（半導體應用)"
      Case 29
         GetItemNm = "核子工程"
      Case 30
         GetItemNm = "電力；發電、配電、變電、電熱"
      Case 31
         GetItemNm = "基本電子電機元件"
      Case 32
         GetItemNm = "半導体技術"
      Case 33
         GetItemNm = "基本電子電路；通信"
      Case 34
         GetItemNm = "資訊"
      Case 35
         GetItemNm = "電子商務"
      Case Else
         GetItemNm = "其他"
   End Select
End Function

'Add By Sindy 2016/2/26
'案件屬性
Private Function GetItemNm2(intItem As Integer) As String
   GetItemNm2 = ""
   Select Case intItem
      Case 1
         GetItemNm2 = "機械"
      Case 2
         GetItemNm2 = "電子電機"
      Case 3
         GetItemNm2 = "化學生醫"
      Case Else
         GetItemNm2 = "其他"
   End Select
End Function

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   strSql = Printer.DeviceName
   SeekPrintL = Printer.Orientation
   For i = 0 To Printers.Count - 1
      Set Printer = Printers(i)
      Combo1.AddItem Printer.DeviceName, j
      j = j + 1
      If Printer.DeviceName = strSql Then
         SeekPrint = i
      End If
   Next i
   
   Set Printer = Printers(SeekPrint)
   Combo1.Text = Combo1.List(SeekPrint)
   
   txt1(0) = Left(strSrvDate(2), 3) & "01"
   txt1(1) = Left(strSrvDate(2), 5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100133 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub

'*************************************************
'  轉成Excel檔案 (公開公報)
'
'*************************************************
Private Sub ExcelSave_2()
On Error GoTo flgErr
   
   'Modify By Sindy 2016/2/26 產業別
   If Option2(0).Value = True Then
      If strStarYM <> strEndYM Then
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月至" & Left(strEndYM, 4) & "年" & Right(strEndYM, 2) & "月（公開公報）產業分類案件市佔分析.xls"
      Else
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月（公開公報）產業分類案件市佔分析.xls"
      End If
   Else '案件屬性
      If strStarYM <> strEndYM Then
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月至" & Left(strEndYM, 4) & "年" & Right(strEndYM, 2) & "月（公開公報）案件屬性案件市佔分析.xls"
      Else
         strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月（公開公報）案件屬性案件市佔分析.xls"
      End If
   End If
   '2016/2/26 END
   
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   xlsSalesPoint.Workbooks.add
   
   Call ExcelSave_Total_2
   Call ExcelSave_FCP_2
   Call ExcelSave_CCP_2
   
   'Modify By Sindy 2018/5/9
   'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   '2018/5/9 END
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   MsgBox "檔案已產生！電子檔位置：" & strFileName
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub ExcelSave_Total_2()
On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
   wksaccrpt114.Name = "Total"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
'   wksaccrpt114.Range("a1:g1").Select
'   With wksaccrpt114.Range("a1:g1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
   wksaccrpt114.Range("c2").Value = "公開公報"
   lngCounter = 3
   
   'Modify By Sindy 2016/2/26
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   '2016/2/26 END
   
   For i = 1 To intRunItem '36 分類項目
      'Modify By Sindy 2016/2/26
      'If i = 36 Then i = 99 '其他
      If Option2(0).Value = True Then '產業別
         If i = intRunItem Then i = 99 '其他
         strKind = GetItemNm(i)
         strConSql = " and TPG17 is not null and TPG17='" & Format(i, "00") & "'"
      Else '案件屬性
         If i = intRunItem Then i = 9 '其他
         strKind = GetItemNm2(i)
         strConSql = " and TPG18 is not null and TPG18='" & Format(i, "0") & "'"
      End If
      '2016/2/26 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql & " and TPG01=PA11(+) and PA09 = '000' and pa23='1'"
      'Add By Sindy 2016/2/26
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql & " and TPG08='將群'"
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql & " and TPG08='冠群國際'"
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql & " and TPG08='連邦國際'"
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql & " and TPG08='聖島國際'"
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql & " and TPG08='理律法律'"
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 " & strConSql & " and TPG08='台灣國際'"
      '2016/2/26 END
      strSql = strSql & ")"
      
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      '2016/2/26 END
      
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
'         dblVal1 = adoRecordset.Fields(0)
'         dblVal2 = adoRecordset.Fields(1)
'         dblVal3 = adoRecordset.Fields(2)
'         dblVal4 = adoRecordset.Fields(3)
         'Modify By Sindy 2016/2/26
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/2/26 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
'      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      'Modify By Sindy 2016/2/26
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j) '3
   '      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1) '4
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
   '      If dblVal(2) = 0 Then
   '         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
   '      Else
   '         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
   '      End If
         j = j + 1
         lngCounter = lngCounter + 1
      Next j
      '2016/2/26 END
      '合計
'      dblValTot_1 = dblValTot_1 + dblVal1
'      dblValTot_2 = dblValTot_2 + dblVal2
'      dblValTot_3 = dblValTot_3 + dblVal3
'      dblValTot_4 = dblValTot_4 + dblVal4
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2016/2/26 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
'   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   'Modify By Sindy 2016/2/26
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j) '3
   '   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1) '4
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
   '   If dblValTot(2) = 0 Then
   '      wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
   '   Else
   '      wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
   '   End If
      j = j + 1
      lngCounter = lngCounter + 1
   Next j
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub ExcelSave_FCP_2()
On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(2)
   wksaccrpt114.Name = "FCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
'   wksaccrpt114.Range("a1:g1").Select
'   With wksaccrpt114.Range("a1:g1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
   wksaccrpt114.Range("c2").Value = "公開公報"
   lngCounter = 3
   
   'Modify By Sindy 2016/2/26
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   '2016/2/26 END
   
   For i = 1 To intRunItem '36 分類項目
      'Modify By Sindy 2016/2/26
      'If i = 36 Then i = 99 '其他
      If Option2(0).Value = True Then '產業別
         If i = intRunItem Then i = 99 '其他
         strKind = GetItemNm(i)
         strConSql = " and TPG17 is not null and TPG17='" & Format(i, "00") & "'"
      Else '案件屬性
         If i = intRunItem Then i = 9 '其他
         strKind = GetItemNm2(i)
         strConSql = " and TPG18 is not null and TPG18='" & Format(i, "0") & "'"
      End If
      '2016/2/26 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql & " and TPG01=PA11(+) and PA09 = '000' and pa23='1'"
      'Add By Sindy 2016/2/26
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql & " and TPG08='將群'"
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql & " and TPG08='冠群國際'"
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql & " and TPG08='連邦國際'"
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql & " and TPG08='聖島國際'"
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql & " and TPG08='理律法律'"
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG06>'010' " & strConSql & " and TPG08='台灣國際'"
      '2016/2/26 END
      strSql = strSql & ")"
      
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      '2016/2/26 END
      
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
'         dblVal1 = adoRecordset.Fields(0)
'         dblVal2 = adoRecordset.Fields(1)
'         dblVal3 = adoRecordset.Fields(2)
'         dblVal4 = adoRecordset.Fields(3)
         'Modify By Sindy 2016/2/26
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/2/26 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
'      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j) '3
   '      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1) '
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
   '      If dblVal(2) = 0 Then
   '         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
   '      Else
   '         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
   '      End If
         j = j + 1
         lngCounter = lngCounter + 1
      Next j
      '2016/2/26 END
      '合計
'      dblValTot_1 = dblValTot_1 + dblVal1
'      dblValTot_2 = dblValTot_2 + dblVal2
'      dblValTot_3 = dblValTot_3 + dblVal3
'      dblValTot_4 = dblValTot_4 + dblVal4
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2016/2/26 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
'   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   'Modify By Sindy 2016/2/26
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
   '   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
   '   If dblValTot(2) = 0 Then
   '      wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
   '   Else
   '      wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
   '   End If
      j = j + 1
      lngCounter = lngCounter + 1
   Next j
   '2016/2/26 END
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub ExcelSave_CCP_2()
On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(3)
   wksaccrpt114.Name = "CCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
'   wksaccrpt114.Range("a1:g1").Select
'   With wksaccrpt114.Range("a1:g1")
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlBottom
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .ShrinkToFit = False
'       .MergeCells = True
'   End With
   wksaccrpt114.Range("c2").Value = "公開公報"
   lngCounter = 3
   
   'Modify By Sindy 2016/2/26
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   '2016/2/26 END
   
   For i = 1 To intRunItem '36 分類項目
      'Modify By Sindy 2016/2/26
      'If i = 36 Then i = 99 '其他
      If Option2(0).Value = True Then '產業別
         If i = intRunItem Then i = 99 '其他
         strKind = GetItemNm(i)
         strConSql = " and TPG17 is not null and TPG17='" & Format(i, "00") & "'"
      Else '案件屬性
         If i = intRunItem Then i = 9 '其他
         strKind = GetItemNm2(i)
         strConSql = " and TPG18 is not null and TPG18='" & Format(i, "0") & "'"
      End If
      '2016/2/26 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql & " and TPG01=PA11(+) and PA09 = '000' and pa23='1'"
      'Add By Sindy 2016/2/26
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql & " and TPG08='將群'"
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql & " and TPG08='冠群國際'"
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql & " and TPG08='連邦國際'"
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql & " and TPG08='聖島國際'"
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql & " and TPG08='理律法律'"
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and (TPG06<='010' or TPG06 is null) " & strConSql & " and TPG08='台灣國際'"
      '2016/2/26 END
      strSql = strSql & ")"
      
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      '2016/2/26 END
      
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
'         dblVal1 = adoRecordset.Fields(0)
'         dblVal2 = adoRecordset.Fields(1)
'         dblVal3 = adoRecordset.Fields(2)
'         dblVal4 = adoRecordset.Fields(3)
         'Modify By Sindy 2016/2/26
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/2/26 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
'      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      'Modify By Sindy 2016/2/26
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
   '      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
   '      If dblVal(2) = 0 Then
   '         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
   '      Else
   '         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
   '      End If
         j = j + 1
         lngCounter = lngCounter + 1
      Next j
      '2016/2/26 END
      '合計
'      dblValTot_1 = dblValTot_1 + dblVal1
'      dblValTot_2 = dblValTot_2 + dblVal2
'      dblValTot_3 = dblValTot_3 + dblVal3
'      dblValTot_4 = dblValTot_4 + dblVal4
      'Modify By Sindy 2016/2/26
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2016/2/26 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
'   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   'Modify By Sindy 2016/2/26
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
   '   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
   '   If dblValTot(2) = 0 Then
   '      wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
   '   Else
   '      wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
   '   End If
      j = j + 1
      lngCounter = lngCounter + 1
   Next j
   '2016/2/26 END
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub
