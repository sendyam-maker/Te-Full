VERSION 5.00
Begin VB.Form frm100130 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報IPC分類案件市佔分析"
   ClientHeight    =   3360
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5320
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5320
   Begin VB.Frame Frame2 
      BorderStyle     =   0  '沒有框線
      Height          =   705
      Left            =   870
      TabIndex        =   12
      Top             =   840
      Width           =   2145
      Begin VB.OptionButton Option1 
         Caption         =   "公開公報"
         Height          =   180
         Index           =   1
         Left            =   240
         TabIndex        =   1
         Top             =   420
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "公報"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Top             =   90
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2970
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1800
      Width           =   825
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   180
      TabIndex        =   7
      Top             =   2700
      Visible         =   0   'False
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   6
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1920
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1800
      Width           =   825
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4215
      TabIndex        =   5
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "產生Excel檔(&E)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "註：IPC分類資料開始於101年01月。"
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   750
      TabIndex        =   11
      Top             =   2490
      Width           =   2880
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "~"
      Height          =   180
      Left            =   2790
      TabIndex        =   10
      Top             =   1830
      Width           =   150
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "起迄年月："
      Height          =   180
      Left            =   990
      TabIndex        =   9
      Top             =   1830
      Width           =   900
   End
End
Attribute VB_Name = "frm100130"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/19 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Create by Sindy 2012/8/20
Option Explicit

Dim strStarYM As String, strEndYM As String
Dim strCountYear As String, strCountMon As String, strChkData As String
Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim lngCounter As Long, i As Integer, j As Integer
Dim strFileName As String, strKind As String
Dim dblVal(1 To 100) As Double 'Modify By Sindy 2014/12/8
Dim dblValTot(1 To 100) As Double 'Modify By Sindy 2014/12/8
Dim intRunCnt As Integer 'Add By Sindy 2014/12/8


Private Sub cmdOK_Click(Index As Integer)
Dim dblCnt As Double
   
   Select Case Index
      Case 0
         ClearQueryLog (Me.Name) 'Add By Sindy 2024/3/22 清除查詢印表記錄檔欄位
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
         
         'Add By Sindy 2024/3/22
         If Option1(1).Value = True Then
            pub_QL05 = pub_QL05 & ";公開公報"
         ElseIf Option1(0).Value = True Then
            pub_QL05 = pub_QL05 & ";公報"
         End If
         pub_QL05 = pub_QL05 & ";起迄年月:" & strStarYM & "~" & strEndYM
         '2024/3/22 END
         
         '逐月檢查有無資料
         strCountYear = Left(strStarYM, 4)
         strCountMon = Right(strStarYM, 2)
         strChkData = "": dblCnt = 0
         Do While Val(strCountYear & strCountMon) <= Val(strEndYM)
            'Modify By Sindy 2017/2/17
            If Option1(1).Value = True Then '公開公報
               strExc(0) = "select count(*) from tpgazette Where TPg03>=" & Val(strCountYear & strCountMon) & "01 and TPg03<=" & Val(strCountYear & strCountMon) & "31" ' and TPg16 is not null 'Modify By Sindy 2024/6/3 mark, 示為其他類
            Else '公報
            '2017/2/17 END
               strExc(0) = "select count(*) from tpbulletin Where TPB03>=" & Val(strCountYear & strCountMon) & "01 and TPB03<=" & Val(strCountYear & strCountMon) & "31" ' and TPB11 is not null 'Modify By Sindy 2024/6/3 mark, 示為其他類
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               dblCnt = dblCnt + RsTemp.Fields(0)
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
         If strChkData <> "" Then
            strChkData = Mid(strChkData, 2)
            If strStarYM = strEndYM Then
               InsertQueryLog (0) 'Add By Sindy 2024/3/22
               MsgBox "查無資料!!!", vbExclamation + vbOKOnly
               txt1(0).SetFocus
               Exit Sub
            Else
               If MsgBox(strChkData & "查無" & IIf(Option1(0).Value = True, "公報", "公開公報") & "資料，確定還要進行統計嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  pub_QL05 = pub_QL05 & ";" & strChkData & "查無" & IIf(Option1(0).Value = True, "公報", "公開公報") & "資料，放棄統計!"
                  InsertQueryLog (dblCnt) 'Add By Sindy 2024/3/22
                  Exit Sub
               End If
            End If
         End If
         
         InsertQueryLog (dblCnt) 'Add By Sindy 2024/3/22
         Screen.MousePointer = vbHourglass
         'Modify By Sindy 2017/2/17
         If Option1(1).Value = True Then '公開公報
            Call ExcelSave_TPG
         Else '公報
         '2017/2/17 END
            Call ExcelSave
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

'*************************************************
'  統計（公報）, 轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave()
On Error GoTo flgErr
   
   intRunCnt = 16
   
   If strStarYM <> strEndYM Then
      strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月至" & Left(strEndYM, 4) & "年" & Right(strEndYM, 2) & "月公報IPC分類案件市佔分析.xls"
   Else
      strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月公報IPC分類案件市佔分析.xls"
   End If
   
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   xlsSalesPoint.SheetsInNewWorkbook = 5 '4 '3 'Add By Sindy 2019/3/12 Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
   xlsSalesPoint.Workbooks.add
   
   Call ExcelSave_Total
   Call ExcelSave_FCP
   Call ExcelSave_FCP_J 'Add By Sindy 2019/8/14
   Call ExcelSave_CCP
   Call ExcelSave_MCP 'Add By Sindy 2021/2/19 再增加工作表MCP，區分大陸來的案件
   
   'Modify By Sindy 2018/3/6
   'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   '2018/3/6 END
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
Dim strSQLCon As String 'Add By Sindy 2024/6/3

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
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   'Modify By Sindy 2014/12/8
   For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
      dblValTot(j) = 0
   Next
   '2014/12/8 END
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPB11 is null or TPB11='12')"
      Else
         strSQLCon = " and TPB11<>'11' and TPB11 is not null and TPB11='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M'" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      'Add By Sindy 2014/12/8
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='將群'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='冠群國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='連邦國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='連邦國際'" & strSQLCon
      '2014/12/8 END
      'Add By Sindy 2014/12/27
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='聖島國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='聖島國際'" & strSQLCon
      '2014/12/27 END
      'Add By Sindy 2016/1/13
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='理律法律'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='台灣國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='台灣國際'" & strSQLCon
      '2016/1/13 END
      strSql = strSql & ")"
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblVal(j) = 0
      Next
      '2014/12/8 END
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         'Modify By Sindy 2016/1/13
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/1/13 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
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
      '合計
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2014/12/8 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
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
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(2)
   wksaccrpt114.Name = "FCP非日本"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   'Modify By Sindy 2014/12/8
   For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
      dblValTot(j) = 0
   Next
   '2014/12/8 END
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPB11 is null or TPB11='12') and (TPB06>010 and TPB06<>'011' and TPB06<>'020')"
      Else
         strSQLCon = " and TPB11<>'11' and TPB11 is not null and (TPB06>010 and TPB06<>'011' and TPB06<>'020') and TPB11='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M'" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      'Add By Sindy 2014/12/8
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='將群'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='冠群國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='連邦國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='連邦國際'" & strSQLCon
      '2014/12/8 END
      'Add By Sindy 2014/12/27
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='聖島國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='聖島國際'" & strSQLCon
      '2014/12/27 END
      'Add By Sindy 2016/1/13
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='理律法律'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='台灣國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='台灣國際'" & strSQLCon
      '2016/1/13 END
      strSql = strSql & ")"
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblVal(j) = 0
      Next
      '2014/12/8 END
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         'Modify By Sindy 2016/1/13
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/1/13 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
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
      '合計
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2014/12/8 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
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
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2019/8/14 專利日本部
Private Sub ExcelSave_FCP_J()
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(3)
   wksaccrpt114.Name = "FCP日本"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   'Modify By Sindy 2014/12/8
   For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
      dblValTot(j) = 0
   Next
   '2014/12/8 END
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPB11 is null or TPB11='12') and TPB06='011'"
      Else
         strSQLCon = " and TPB11<>'11' and TPB11 is not null and TPB06='011' and TPB11='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M'" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      'Add By Sindy 2014/12/8
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='將群'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='冠群國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='連邦國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='連邦國際'" & strSQLCon
      '2014/12/8 END
      'Add By Sindy 2014/12/27
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='聖島國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='聖島國際'" & strSQLCon
      '2014/12/27 END
      'Add By Sindy 2016/1/13
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='理律法律'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='台灣國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='台灣國際'" & strSQLCon
      '2016/1/13 END
      strSql = strSql & ")"
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblVal(j) = 0
      Next
      '2014/12/8 END
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         'Modify By Sindy 2016/1/13
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/1/13 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
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
      '合計
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2014/12/8 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
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
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2021/2/19 區分大陸來的案件
Private Sub ExcelSave_MCP()
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(5)
   wksaccrpt114.Name = "MCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   'Modify By Sindy 2014/12/8
   For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
      dblValTot(j) = 0
   Next
   '2014/12/8 END
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPB11 is null or TPB11='12') and TPB06='020'"
      Else
         strSQLCon = " and TPB11<>'11' and TPB11 is not null and TPB06='020' and TPB11='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M'" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      'Add By Sindy 2014/12/8
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='將群'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='冠群國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='連邦國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='連邦國際'" & strSQLCon
      '2014/12/8 END
      'Add By Sindy 2014/12/27
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='聖島國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='聖島國際'" & strSQLCon
      '2014/12/27 END
      'Add By Sindy 2016/1/13
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='理律法律'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='台灣國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='台灣國際'" & strSQLCon
      '2016/1/13 END
      strSql = strSql & ")"
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblVal(j) = 0
      Next
      '2014/12/8 END
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         'Modify By Sindy 2016/1/13
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/1/13 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
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
      '合計
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2014/12/8 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
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
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub ExcelSave_CCP()
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(4)
   wksaccrpt114.Name = "CCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   'Modify By Sindy 2014/12/8
   For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
      dblValTot(j) = 0
   Next
   '2014/12/8 END
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPB11 is null or TPB11='12') and (TPB06<=010 or TPB06 is null)"
      Else
         strSQLCon = " and TPB11<>'11' and TPB11 is not null and (TPB06<=010 or TPB06 is null) and TPB11='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8),sum(t9),sum(t10),sum(t11),sum(t12),sum(t13),sum(t14),sum(t15),sum(t16) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M'" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 FROM Patent,tpbulletin WHERE TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB01=PA11(+) and PA09 = '000' and pa23='1' and TPB08='台一國際'" & strSQLCon
      'Add By Sindy 2014/12/8
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='將群'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='冠群國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,count(*) as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='連邦國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,count(*) as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='連邦國際'" & strSQLCon
      '2014/12/8 END
      'Add By Sindy 2014/12/27
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,count(*) as t11,0 as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='聖島國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,count(*) as t12,0 as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='聖島國際'" & strSQLCon
      '2014/12/27 END
      'Add By Sindy 2016/1/13
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,count(*) as t13,0 as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='理律法律'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,count(*) as t14,0 as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,count(*) as t15,0 as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='I' and TPB08='台灣國際'" & strSQLCon
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8,0 as t9,0 as t10,0 as t11,0 as t12,0 as t13,0 as t14,0 as t15,count(*) as t16 from tpbulletin where TPB03>=" & strStarYM & "01 and TPB03<=" & strEndYM & "31 and substr(TPB02,1,1)='M' and TPB08='台灣國際'" & strSQLCon
      '2016/1/13 END
      strSql = strSql & ")"
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblVal(j) = 0
      Next
      '2014/12/8 END
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         'Modify By Sindy 2016/1/13
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
         '2016/1/13 END
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 3 To intRunCnt
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
         If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
         If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
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
      '合計
      'Modify By Sindy 2014/12/8
      For j = 1 To intRunCnt '12 '10 Modify By Sindy 2014/12/27
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
      '2014/12/8 END
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 3 To intRunCnt
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 9 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 11 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島" 'Add By Sindy 2014/12/27
      If j = 13 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律" 'Add By Sindy 2016/1/13
      If j = 15 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣" 'Add By Sindy 2016/1/13
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
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   'Modified by Lydia 2019/10/30 改成共用模組
'   strSql = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
'   For i = 0 To Printers.Count - 1
'      Set Printer = Printers(i)
'      Combo1.AddItem Printer.DeviceName, j
'      j = j + 1
'      If Printer.DeviceName = strSql Then
'         SeekPrint = i
'      End If
'   Next i
'
'   Set Printer = Printers(SeekPrint)
'   Combo1.Text = Combo1.List(SeekPrint)
   PUB_SetPrinter Me.Name, Combo1
   
   txt1(0) = Left(strSrvDate(2), 3) & "01"
   txt1(1) = Left(strSrvDate(2), 5)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100130 = Nothing
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

'Add By Sindy 2017/2/17
'*************************************************
'  統計（公開公報）, 轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave_TPG()
On Error GoTo flgErr
   
   intRunCnt = 8
   
   If strStarYM <> strEndYM Then
      strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月至" & Left(strEndYM, 4) & "年" & Right(strEndYM, 2) & "月公開公報IPC分類案件市佔分析.xls"
   Else
      strFileName = PUB_Getdesktop & "\" & Left(strStarYM, 4) & "年" & Right(strStarYM, 2) & "月公開公報IPC分類案件市佔分析.xls"
   End If
   
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   xlsSalesPoint.SheetsInNewWorkbook = 5 '4 '3 'Add By Sindy 2019/3/12 Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
   xlsSalesPoint.Workbooks.add
   
   Call ExcelSave_Total_TPG
   Call ExcelSave_FCP_TPG
   Call ExcelSave_FCP_TPG_J 'Add By Sindy 2019/8/14
   Call ExcelSave_CCP_TPG
   Call ExcelSave_MCP_TPG 'Add By Sindy 2021/2/19 再增加工作表MCP，區分大陸來的案件
   
   'Modify By Sindy 2018/3/6
   'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   '2018/3/6 END
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

'Add By Sindy 2017/2/17
Private Sub ExcelSave_Total_TPG()
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(1)
   wksaccrpt114.Name = "Total"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   'wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   'wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPG16 is null or TPG16='12')"
      Else
         strSQLCon = " and TPG16<>'11' and TPG16 is not null and TPG16='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG01=PA11(+) and PA09 = '000' and pa23='1' and TPG08='台一國際'" & strSQLCon
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='連邦國際'" & strSQLCon
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='聖島國際'" & strSQLCon
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='台灣國際'" & strSQLCon
      strSql = strSql & ")"
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 2 To intRunCnt
         If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
'         If dblVal(2) = 0 Then
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'         Else
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
'         End If
         lngCounter = lngCounter + 1
      Next j
      '合計
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 2 To intRunCnt
      If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
'      If dblValTot(2) = 0 Then
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'      Else
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
'      End If
      lngCounter = lngCounter + 1
   Next j
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2017/2/17
Private Sub ExcelSave_FCP_TPG()
Dim strSQLCon As String 'Add By Sindy 2024/6/3
   
On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(2)
   wksaccrpt114.Name = "FCP非日本"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   'wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   'wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPG16 is null or TPG16='12') and (TPG06>010 and TPG06<>'011' and TPG06<>'020')"
      Else
         strSQLCon = " and TPG16<>'11' and TPG16 is not null and (TPG06>010 and TPG06<>'011' and TPG06<>'020') and TPG16='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG01=PA11(+) and PA09 = '000' and pa23='1' and TPG08='台一國際'" & strSQLCon
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='連邦國際'" & strSQLCon
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='聖島國際'" & strSQLCon
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='台灣國際'" & strSQLCon
      strSql = strSql & ")"
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 2 To intRunCnt
         If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
'         If dblVal(2) = 0 Then
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'         Else
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
'         End If
         lngCounter = lngCounter + 1
      Next j
      '合計
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 2 To intRunCnt
      If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
'      If dblValTot(2) = 0 Then
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'      Else
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
'      End If
      lngCounter = lngCounter + 1
   Next j
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2019/8/14 專利日本部
Private Sub ExcelSave_FCP_TPG_J()
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(3)
   wksaccrpt114.Name = "FCP日本"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   'wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   'wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPG16 is null or TPG16='12') and TPG06='011'"
      Else
         strSQLCon = " and TPG16<>'11' and TPG16 is not null and TPG06='011' and TPG16='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG01=PA11(+) and PA09 = '000' and pa23='1' and TPG08='台一國際'" & strSQLCon
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='連邦國際'" & strSQLCon
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='聖島國際'" & strSQLCon
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='台灣國際'" & strSQLCon
      strSql = strSql & ")"
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 2 To intRunCnt
         If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
'         If dblVal(2) = 0 Then
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'         Else
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
'         End If
         lngCounter = lngCounter + 1
      Next j
      '合計
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 2 To intRunCnt
      If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
'      If dblValTot(2) = 0 Then
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'      Else
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
'      End If
      lngCounter = lngCounter + 1
   Next j
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2021/2/19 區分大陸來的案件
Private Sub ExcelSave_MCP_TPG()
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(5)
   wksaccrpt114.Name = "MCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   'wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   'wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPG16 is null or TPG16='12') and TPG06='020'"
      Else
         strSQLCon = " and TPG16<>'11' and TPG16 is not null and TPG06='020' and TPG16='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG01=PA11(+) and PA09 = '000' and pa23='1' and TPG08='台一國際'" & strSQLCon
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='連邦國際'" & strSQLCon
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='聖島國際'" & strSQLCon
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='台灣國際'" & strSQLCon
      strSql = strSql & ")"
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 2 To intRunCnt
         If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
'         If dblVal(2) = 0 Then
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'         Else
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
'         End If
         lngCounter = lngCounter + 1
      Next j
      '合計
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 2 To intRunCnt
      If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
'      If dblValTot(2) = 0 Then
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'      Else
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
'      End If
      lngCounter = lngCounter + 1
   Next j
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

'Add By Sindy 2017/2/17
Private Sub ExcelSave_CCP_TPG()
Dim strSQLCon As String 'Add By Sindy 2024/6/3

On Error GoTo flgErr
   
   Set wksaccrpt114 = xlsSalesPoint.Worksheets(4)
   wksaccrpt114.Name = "CCP"
   wksaccrpt114.Columns("a:a").ColumnWidth = 13
   wksaccrpt114.Columns("b:b").ColumnWidth = 13
   wksaccrpt114.Columns("c:c").ColumnWidth = 13
   'wksaccrpt114.Columns("d:d").ColumnWidth = 13
   wksaccrpt114.Columns("e:e").ColumnWidth = 13
   wksaccrpt114.Columns("f:f").ColumnWidth = 13
   wksaccrpt114.Columns("g:g").ColumnWidth = 13
   wksaccrpt114.Range("c1").Value = Val(txt1(0)) + 191100 & "~" & Val(txt1(1)) + 191100
   wksaccrpt114.Range("c2").Value = "發明"
   'wksaccrpt114.Range("d2").Value = "新型"
   lngCounter = 3
   For j = 1 To intRunCnt
      dblValTot(j) = 0
   Next
   For i = 1 To 10 '9
      If i = 1 Then strKind = "半導體"
      If i = 2 Then strKind = "資訊類"
      If i = 3 Then strKind = "通訊類"
      If i = 4 Then strKind = "電力,量測,光"
      If i = 5 Then strKind = "生技"
      If i = 6 Then strKind = "化學"
      If i = 7 Then strKind = "光電"
      If i = 8 Then strKind = "機械"
      If i = 9 Then strKind = "日用品/醫工類"
      'Add By Sindy 2024/6/3
      If i = 10 Then
         strKind = "其他"
         strSQLCon = " and (TPG16 is null or TPG16='12') and (TPG06<=010 or TPG06 is null)"
      Else
         strSQLCon = " and TPG16<>'11' and TPG16 is not null and (TPG06<=010 or TPG06 is null) and TPG16='" & Format(i, "00") & "'"
      End If
      '2024/6/3 END
      
      strSql = "select sum(t1),sum(t2),sum(t3),sum(t4),sum(t5),sum(t6),sum(t7),sum(t8) from("
      '全國
      strSql = strSql & " select count(*) as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31" & strSQLCon
      '台一
      strSql = strSql & " Union"
      strSql = strSql & " SELECT 0 as t1,count(*) as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 FROM Patent,tpgazette WHERE TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG01=PA11(+) and PA09 = '000' and pa23='1' and TPG08='台一國際'" & strSQLCon
      '將群
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,count(*) as t3,0 as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='將群'" & strSQLCon
      '冠群國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,count(*) as t4,0 as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='冠群國際'" & strSQLCon
      '連邦國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,count(*) as t5,0 as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='連邦國際'" & strSQLCon
      '聖島國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,count(*) as t6,0 as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='聖島國際'" & strSQLCon
      '理律法律
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,count(*) as t7,0 as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='理律法律'" & strSQLCon
      '台灣國際
      strSql = strSql & " Union"
      strSql = strSql & " select 0 as t1,0 as t2,0 as t3,0 as t4,0 as t5,0 as t6,0 as t7,count(*) as t8 from tpgazette where TPG03>=" & strStarYM & "01 and TPG03<=" & strEndYM & "31 and TPG08='台灣國際'" & strSQLCon
      strSql = strSql & ")"
      For j = 1 To intRunCnt
         dblVal(j) = 0
      Next
      If adoRecordset.State = adStateOpen Then
         adoRecordset.Close
      End If
      adoRecordset.CursorLocation = adUseClient
      adoRecordset.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoRecordset.RecordCount > 0 Then
         For j = 1 To intRunCnt
            dblVal(j) = adoRecordset.Fields(j - 1)
         Next
      End If
      wksaccrpt114.Range("a" & lngCounter).Value = strKind
      wksaccrpt114.Range("b" & lngCounter).Value = "全國"
      wksaccrpt114.Range("c" & lngCounter).Value = dblVal(1)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(2)
      lngCounter = lngCounter + 1
      For j = 2 To intRunCnt
         If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
         If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
         If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
         If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
         If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
         If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
         If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
         wksaccrpt114.Range("c" & lngCounter).Value = dblVal(j)
         'wksaccrpt114.Range("d" & lngCounter).Value = dblVal(j + 1)
         lngCounter = lngCounter + 1
         wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
         If dblVal(1) = 0 Then
            wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
         Else
            wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblVal(j) / dblVal(1)) * 100, 3), "#0.00") & "%"
         End If
'         If dblVal(2) = 0 Then
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'         Else
'            wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblVal(j + 1) / dblVal(2)) * 100, 3), "#0.00") & "%"
'         End If
         lngCounter = lngCounter + 1
      Next j
      '合計
      For j = 1 To intRunCnt
         dblValTot(j) = dblValTot(j) + dblVal(j)
      Next
   Next i
   '填入合計
   wksaccrpt114.Range("a" & lngCounter).Value = "合計"
   wksaccrpt114.Range("b" & lngCounter).Value = "全國"
   wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(1)
   'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(2)
   lngCounter = lngCounter + 1
   For j = 2 To intRunCnt
      If j = 2 Then wksaccrpt114.Range("b" & lngCounter).Value = "台一"
      If j = 3 Then wksaccrpt114.Range("b" & lngCounter).Value = "將群"
      If j = 4 Then wksaccrpt114.Range("b" & lngCounter).Value = "冠群"
      If j = 5 Then wksaccrpt114.Range("b" & lngCounter).Value = "連邦"
      If j = 6 Then wksaccrpt114.Range("b" & lngCounter).Value = "聖島"
      If j = 7 Then wksaccrpt114.Range("b" & lngCounter).Value = "理律"
      If j = 8 Then wksaccrpt114.Range("b" & lngCounter).Value = "台灣"
      wksaccrpt114.Range("c" & lngCounter).Value = dblValTot(j)
      'wksaccrpt114.Range("d" & lngCounter).Value = dblValTot(j + 1)
      lngCounter = lngCounter + 1
      wksaccrpt114.Range("b" & lngCounter).Value = "所佔比率"
      If dblValTot(1) = 0 Then
         wksaccrpt114.Range("c" & lngCounter).Value = Format(0, "#0.00") & "%"
      Else
         wksaccrpt114.Range("c" & lngCounter).Value = Format(Round((dblValTot(j) / dblValTot(1)) * 100, 3), "#0.00") & "%"
      End If
'      If dblValTot(2) = 0 Then
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(0, "#0.00") & "%"
'      Else
'         wksaccrpt114.Range("d" & lngCounter).Value = Format(Round((dblValTot(j + 1) / dblValTot(2)) * 100, 3), "#0.00") & "%"
'      End If
      lngCounter = lngCounter + 1
   Next j
   
   Exit Sub
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub
