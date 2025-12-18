VERSION 5.00
Begin VB.Form frm210150 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權部工作報告"
   ClientHeight    =   945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   945
   ScaleWidth      =   4155
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   3195
      Style           =   1  '圖片外觀
      TabIndex        =   1
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Excel"
      Default         =   -1  'True
      Height          =   400
      Left            =   2250
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox txtCloseDate 
      Height          =   285
      Left            =   930
      MaxLength       =   5
      TabIndex        =   0
      Top             =   540
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "(年月)"
      Height          =   180
      Left            =   1935
      TabIndex        =   4
      Top             =   600
      Width           =   585
   End
   Begin VB.Label Label2 
      Caption         =   "統計年月"
      Height          =   180
      Left            =   135
      TabIndex        =   2
      Top             =   585
      Width           =   800
   End
End
Attribute VB_Name = "frm210150"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/25 Form2.0已檢查 (無需修改的物件)
'Create by Amy 2015/04/28
Option Explicit

Dim adoMon As New ADODB.Recordset, adoPMon As New ADODB.Recordset
Dim strMon As String
Dim strRowN1()
Dim strColN(), intWidth()
Dim strFileN As String
Dim xlsCustPoint As New Excel.Application, wksrpt As New Worksheet, intXlsSheet As Integer, bolOpenXls As Boolean
Dim intCounter As Integer, TitleRow As Integer, intField As Integer
Dim ii As Integer, jj As Integer
Dim strWkName As String 'Add by Amy 2018/04/16 for 2010 工作表名稱為中文
Dim strDate1 As String, StrDate2 As String 'Add by Amy 2020/10/19
    
Private Sub Form_Load()

    Dim stST05 As String, stST15 As String
    
    MoveFormToCenter Me
   
    stST15 = PUB_GetStaffST15(strUserNum, 1)
    stST05 = PUB_GetST05(strUserNum)
   
    txtCloseDate = (CompDate(1, -1, strSrvDate(2)) - 19110000) \ 100
     
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function FormCheck() As Boolean
    Dim bCancel As Boolean
    'Add by Amy 2018/05/17
    Dim strA0b01_1 As String, strA0b01_J As String, strA0b05 As String
    
    If txtCloseDate = MsgText(601) Then
        MsgBox Label2 & "不可為空!"
        Exit Function
    End If
    txtCloseDate_Validate bCancel
    If bCancel = True Then Exit Function
    
    'Add by Amy 2018/05/17 增加財務處關閉判斷
    strA0b01_1 = Left(GetA0b01(strA0b05, "1"), 5)
    strA0b01_J = Left(GetA0b01(strA0b05, "J"), 5)
    If Val(txtCloseDate) > Val(strA0b01_1) Or Val(txtCloseDate) > Val(strA0b01_J) Then
        If Val(txtCloseDate) > Val(strA0b01_J) Then strA0b01_1 = strA0b01_J
        If Val(Right(strA0b01_1, 2)) = 12 Then
            MsgBox Label2 & "點數結算期" & Val(Left(strA0b01_1, IIf(Len(strA0b01_1) = 5, 3, 2))) + 1 & "年1月" & _
                                     "財務處尚未確認，故不可操作"
        Else
            MsgBox Label2 & "點數結算期" & Left(strA0b01_1, IIf(Len(strA0b01_1) = 5, 3, 2)) & "年" & Val(Right(strA0b01_1, 2)) + 1 & "月" & _
                                     "財務處尚未確認，故不可操作"
        End If
        txtCloseDate.SetFocus
        txtCloseDate_GotFocus
        Exit Function
    End If
    'end 2018/05/17
    FormCheck = True
End Function

Private Sub cmdExcel_Click()
    If FormCheck = False Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    bolOpenXls = False: intXlsSheet = 1: TitleRow = 1: intField = 65
    strDate1 = "": StrDate2 = "" 'Add by Amy 2020/10/19
    
    If doQuery1 = False Then
        MsgBox "無工作報告資料"
    End If
    If doQuery2 = True Then
        'MsgBox "Excel 已產生~"
    End If
    If bolOpenXls = True Then fn_EndExcel
    Screen.MousePointer = vbDefault
End Sub

Private Sub fn_CreateExcel()
    'Mark by Amy 2015/05/18 不存檔改至顯示於畫面上
'    strFileN = "智權部工作報告" & txtCloseDate & ACDate(ServerDate) & ServerTime & MsgText(43)
'    If Dir(strExcelPath & strFileN) = MsgText(601) Then
'        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'            MkDir strExcelPath
'        End If
'    Else
'        Kill strExcelPath & strFileN
'    End If
    Set xlsCustPoint = New Excel.Application
    xlsCustPoint.Visible = True
    xlsCustPoint.SheetsInNewWorkbook = 2 'Added by Lydia 2019/03/13 預設工作表數量
    xlsCustPoint.Workbooks.add
    bolOpenXls = True
End Sub

Private Sub fn_EndExcel()
    'Mark by Amy 2015/05/18 不存檔改至顯示於畫面上
'    xlsCustPoint.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN
'    xlsCustPoint.Workbooks.Close
'    xlsCustPoint.Quit
    xlsCustPoint.WindowState = wdWindowStateMaximize
    Set xlsCustPoint = Nothing
    Set wksrpt = Nothing
End Sub

'Add by Amy 2020/10/19 部份語法改抓共用Function
Private Function doQuery1() As Boolean
    Dim stVTB0 As String, stVTB1(1) As String, stVTB2(1) As String, stVTB3(1) As String, stVTB11(1) As String, stVTB21(1) As String, stVTB31(1) As String
    Dim strWhere As String
    
On Error GoTo ErrHnd
    
    strDate1 = Val(txtCloseDate)
    If Val(Right(strDate1, 2)) = 1 Then
        If Len(strDate1) = 5 Then
            StrDate2 = Val(Left(txtCloseDate, 3)) - 1 & "12"
        Else
            StrDate2 = Val(Left(txtCloseDate, 2)) - 1 & "12"
        End If
    Else
        StrDate2 = Val(txtCloseDate) - 1
    End If
    strWhere = " And A0205 >= " & StrDate2 & "01 And A0205 <= " & strDate1 & "31"
    
    '人員(避免有業績沒目標或有目標沒業績,造成資料不對,所以都抓)
    stVTB0 = " Select Distinct st01 ID From Staff,PerFormance Where SubStr(st15,1,1)<>'F' And PE01(+)=ST01 And PE02(+)='TOT'" & _
            " And PE03>=" & Val(txtCloseDate) + 191099 & " And  PE03<=" & Val(txtCloseDate) + 191100 & " Group by st01 " & _
          "Union Select Distinct ax209 ID From acc020,acc021,staff Where ax201(+) = a0201 And ax202(+) = a0202" & strWhere & _
            " And st01(+)=ax209 And SubStr(ST15,1,1)<>'F' And (SubStr(ax205,1,1)= '4' Or SubStr(ax205,1,2)='71') " & _
            " And ax209 is not null Group by ax209"
    
    '上月保留:畫面條件當月4191+4192+4194 「貸方」
    stVTB1(0) = GetPoint(1.1, Val(strDate1), Val(strDate1), , , , , Me.Name, True)
    stVTB11(0) = GetPoint(1.2, Val(strDate1), Val(strDate1), , , , , Me.Name, True)
    
    '本月保留:畫面條件當月4191+4192+4194 「借方」
    stVTB2(0) = GetPoint(1.5, Val(strDate1), Val(strDate1), , , , , Me.Name, True)
    stVTB21(0) = GetPoint(1.6, Val(strDate1), Val(strDate1), , , , , Me.Name, True)
   
    '當月實績
    stVTB3(0) = GetPoint(1.3, Val(strDate1), Val(strDate1), , , , , Me.Name, True)
    '當月結餘
    stVTB31(0) = GetPoint(1.4, Val(strDate1), Val(strDate1), , , , , Me.Name, True)
  
    '畫面條件月份-1(上個月 for 本月較上月使用)
    '上月保留 4191+4192+4194 「貸方」
    stVTB1(1) = GetPoint(1.1, Val(StrDate2), Val(StrDate2), , , , , Me.Name, True)
    stVTB11(1) = GetPoint(1.2, Val(StrDate2), Val(StrDate2), , , , , Me.Name, True)
    
    '本月保留:畫面條件當月4191+4192+4194 「借方」
    stVTB2(1) = GetPoint(1.5, Val(StrDate2), Val(StrDate2), , , , , Me.Name, True)
    stVTB21(1) = GetPoint(1.6, Val(StrDate2), Val(StrDate2), , , , , Me.Name, True)
    
    '當月實績
    stVTB3(1) = GetPoint(1.3, Val(StrDate2), Val(StrDate2), , , , , Me.Name, True)
    '當月結餘
    stVTB31(1) = GetPoint(1.4, Val(StrDate2), Val(StrDate2), , , , , Me.Name, True)
    
    'Modify by Amy 2021/05/27 原:SubStr(st15,1,1)<>'F' 11004月,因其他包含L的資料,故將其他拿掉(只顯示智權部)-簡協理
    strMon = "Select ST15,Min(a0902) DepN,(Sum(Nvl(a1.V11,0))+Sum(Nvl(a2.V21,0)))/1000 PreKeep,(Sum(Nvl(a3.v51,0))+Sum(Nvl(a4.v61,0)))/1000 NowKeep,Sum(Nvl(a5.V31,0))/1000 NowA,'' NowR,Sum(Nvl(a6.V41,0))/1000 NowS" & _
                                                                    ",(Sum(Nvl(b1.V11,0))+Sum(Nvl(b2.V21,0)))/1000 PreM1,(Sum(Nvl(b3.V51,0))+Sum(Nvl(b4.V61,0)))/1000 PreM2,Sum(Nvl(b5.V31,0))/1000 PreM3,Sum(Nvl(b6.V41,0))/1000 PreM4" & _
        " From (" & stVTB0 & ") V0,(" & stVTB1(0) & ") a1,(" & stVTB11(0) & ") a2,(" & stVTB2(0) & ") a3,(" & stVTB21(0) & ") a4,(" & stVTB3(0) & ") a5,(" & stVTB31(0) & ") a6" & _
                               ",Staff,Acc090,(" & stVTB1(1) & ") b1,(" & stVTB11(1) & ") b2,(" & stVTB2(1) & ") b3,(" & stVTB21(1) & ") b4,(" & stVTB3(1) & ") b5,(" & stVTB31(1) & ") b6" & _
        " Where a1.V10(+)=ID And a2.V20(+)=ID And a3.V50(+)=ID And a4.V60(+)=ID And a5.V30(+)=ID And a6.V40(+)=ID And st01(+)=ID " & _
            "And b1.V10(+)=ID And b2.V20(+)=ID And b3.V50(+)=ID And b4.V60(+)=ID And b5.V30(+)=ID And b6.V40(+)=ID And a0901(+)=st15 And SubStr(st15,1,1)='S' " & _
           " And (a1.V11>0 or a2.V21>0 or a3.V51>0 or a4.V61>0 or a5.V31>0 or a6.V41>0 or b1.V11>0 or b2.V21>0 or b3.V51>0 or b4.V61>0 or b5.V31>0 or b6.V41>0 ) " & _
        "Group by st15"
    
    If adoMon.State = adStateOpen Then adoMon.Close
    adoMon.CursorLocation = adUseClient
    adoMon.Open strMon, adoTaie, adOpenStatic, adLockReadOnly
    If adoMon.RecordCount > 0 Then
        fn_CreateExcel
        intCounter = 1
        If strWkName = MsgText(601) Then strWkName = Left(xlsCustPoint.Worksheets(1).Name, Len(xlsCustPoint.Worksheets(1).Name) - 1)
        Set wksrpt = xlsCustPoint.Worksheets(strWkName & intXlsSheet)
        SaveExcel1
        intXlsSheet = intXlsSheet + 1
        doQuery1 = True
    End If
    adoMon.Close
    Exit Function
    
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Mark by Amy 2020/10/19
Private Function doQuery1_Old() As Boolean
'    Dim stVTB0 As String, stVTB1(1) As String, stVTB2(1) As String, stVTB3(1) As String
'    Dim strWhere(2) As String
'
'On Error GoTo ErrHnd
'
'    strWhere(0) = " And A0205 >= " & Val(txtCloseDate) & "01 And A0205 <= " & Val(txtCloseDate) & "31"
'    If Val(Right(txtCloseDate, 2)) = 1 Then
'        If Len(txtCloseDate) = 5 Then
'            strWhere(1) = " And A0205 >= " & Val(Left(txtCloseDate, 3)) - 1 & "1201 And A0205 <= " & Val(Left(txtCloseDate, 3)) - 1 & "1231"
'            strWhere(2) = " And A0205 >= " & Val(Left(txtCloseDate, 3)) - 1 & "1201 And A0205 <= " & Val(txtCloseDate) & "31"
'        Else
'            strWhere(1) = " And A0205 >= " & Val(Left(txtCloseDate, 2)) - 1 & "1201 And A0205 <= " & Val(Left(txtCloseDate, 2)) - 1 & "1231"
'            strWhere(2) = " And A0205 >= " & Val(Left(txtCloseDate, 2)) - 1 & "1201 And A0205 <= " & Val(txtCloseDate) & "31"
'        End If
'    Else
'        strWhere(1) = " And A0205 >= " & Val(txtCloseDate) - 1 & "01 And A0205 <= " & Val(txtCloseDate) - 1 & "31"
'        strWhere(2) = " And A0205 >= " & Val(txtCloseDate) - 1 & "01 And A0205 <= " & Val(txtCloseDate) & "31"
'    End If
'
'    '人員(避免有業績沒目標或有目標沒業績,造成資料不對,所以都抓)
'    stVTB0 = " Select Distinct st01 ID From Staff,PerFormance Where SubStr(st15,1,1)<>'F' And PE01(+)=ST01 And PE02(+)='TOT'" & _
'        " And PE03>=" & Val(txtCloseDate) + 191099 & " And  PE03<=" & Val(txtCloseDate) + 191100 & "Group by st01 " & _
'    "Union Select Distinct ax209 ID From acc020,acc021,staff Where ax201(+) = a0201 And ax202(+) = a0202" & strWhere(2) & _
'        " And st01(+)=ax209 And SubStr(ST15,1,1)<>'F' And (SubStr(ax205,1,2)= '41' Or SubStr(ax205,1,2)='71') " & _
'        " And ax209 is not null Group by ax209"
'    '上月未報:畫面條件當月4191+4192+4194 「貸方」
'    'Modify by Amy 2016/11/17 排除有轉撥
'    stVTB1(0) = "Select ax209 V10, Sum(ax207) V11" & _
'        " From acc020, acc021,Staff" & _
'        " Where ax201(+) = a0201  And ax202(+) = a0202" & strWhere(0) & _
'        " And st01(+)=ax209 And SubStr(ST15,1,1)<>'F' And InStr(ax212,'轉撥')=0 " & _
'        " And (ax205= '4191' Or ax205='4192' or ax205='4194') Group by ax209"
'    '本月未報:畫面條件當月4191+4192+4194 「借方」
'    'Modify by Amy 2016/11/17 排除有轉撥
'    stVTB2(0) = "Select ax209 V20, Sum(ax206) V21" & _
'        " From acc020, acc021,Staff" & _
'        " Where ax201(+) = a0201  And ax202(+) = a0202" & strWhere(0) & _
'        " And st01(+)=ax209 And ST01>'60' And ST01<'F' And SubStr(ST15,1,1)<>'F' And InStr(ax212,'轉撥')=0 " & _
'        " And (ax205= '4191' Or ax205='4192' or ax205='4194') Group by ax209"
'    '本月實際達成=當月實績+當月結餘
'    'Modify by Amy 2019/10/23 增加創新業務組用收入 420101 原:SubStr(ax205, 1, 2) = '41'
'    stVTB3(0) = "Select ax209 V30, Sum(ax207-ax206) V31" & _
'        " From acc020, acc021,Staff" & _
'        " Where ax201(+) = a0201  And ax202(+) = a0202" & strWhere(0) & _
'        " And ST01(+)=ax209 And SubStr(ST15,1,1)<>'F' " & _
'        " And (SubStr(ax205, 1, 1) = '4' Or ax205='7121') And ax209 is not null And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
'        " Group by ax209"
'
'    '畫面條件月份-1
'    '上月未報 4191+4192+4194 「貸方」
'    'Modify by Amy 2016/11/17 排除有轉撥
'    stVTB1(1) = "Select ax209 V40, Sum(ax207) V41" & _
'        " From acc020, acc021,Staff" & _
'        " Where ax201(+) = a0201  And ax202(+) = a0202" & strWhere(1) & _
'        " And st01(+)=ax209 And ST01>'60' And ST01<'F' And SubStr(ST15,1,1)<>'F' And InStr(ax212,'轉撥')=0 " & _
'        " And (ax205= '4191' Or ax205='4192' or ax205='4194') Group by ax209"
'    '本月未報:畫面條件當月4191+4192+4194 「借方」
'    'Modify by Amy 2016/11/17 排除有轉撥
'    stVTB2(1) = "Select ax209 V50, Sum(ax206) V51" & _
'        " From acc020, acc021,Staff" & _
'        " Where ax201(+) = a0201  And ax202(+) = a0202" & strWhere(1) & _
'        " And st01(+)=ax209 And SubStr(ST15,1,1)<>'F' And InStr(ax212,'轉撥')=0 " & _
'        " And (ax205= '4191' Or ax205='4192' or ax205='4194') Group by ax209"
'    '上月實際達成
'    'Modify by Amy 2019/10/23 增加創新業務組用收入 420101 原:SubStr(ax205, 1, 2) = '41'
'    stVTB3(1) = "Select ax209 V60, Sum(ax207-ax206) V61" & _
'        " From acc020, acc021,Staff" & _
'        " Where ax201(+) = a0201  And ax202(+) = a0202" & strWhere(1) & _
'        " And ST01(+)=ax209 And SubStr(ST15,1,1)<>'F' " & _
'        " And (SubStr(ax205, 1,1) = '4' Or ax205='7121') And ax209 is not null And Not( ax205='4191' or ax205='4192' or ax205='4194')" & _
'        " Group by ax209"
'
'    strMon = "Select ST15,Min(a0902) DepN,(NVL(Sum(V11),0))/1000 C1,NVL(Sum(V21),0)/1000 C2" & _
'        ",NVL(Sum(V31),0)/1000 C3,NVL(Sum(V41),0)/1000 C4,NVL(Sum(V51),0)/1000 C5,NVL(Sum(V61),0)/1000 C6" & _
'        " From (" & stVTB0 & ") V0,(" & stVTB1(0) & ") V1,(" & stVTB2(0) & ") V2,(" & stVTB3(0) & ") V3,(" & stVTB1(1) & ") V4," & _
'        "(" & stVTB2(1) & ") V5,(" & stVTB3(1) & ") V6,Staff,Acc090" & _
'        " Where V10(+)=ID And V20(+)=ID And V30(+)=ID And V40(+)=ID And V50(+)=ID And V60(+)=ID And st01(+)=ID And a0901(+)=st15 " & _
'        " And (V11>0 or V21>0 or V31>0 or V41>0 or V51>0 or V61>0) Group by st15"
'
'    If adoMon.State = adStateOpen Then adoMon.Close
'    adoMon.CursorLocation = adUseClient
'    adoMon.Open strMon, adoTaie, adOpenStatic, adLockReadOnly
'    If adoMon.RecordCount > 0 Then
'        fn_CreateExcel
'        intCounter = 1
'        'Modify by Amy 2018/04/16 for 工作表名稱改為中文
'        If strWkName = MsgText(601) Then strWkName = Left(xlsCustPoint.Worksheets(1).Name, Len(xlsCustPoint.Worksheets(1).Name) - 1)
'        Set wksrpt = xlsCustPoint.Worksheets(strWkName & intXlsSheet)
'        'end 2018/04/16
'        SaveExcel1
'        intXlsSheet = intXlsSheet + 1
'        doQuery1 = True
'    End If
'    adoMon.Close
'    Exit Function
'
'ErrHnd:
'   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function
'end 2020/10/19

Private Sub SaveExcel1()
    Dim strTp As String, intTp As Integer
    Dim StartRow As Integer, EndRow As Integer
    'Modify by Amy 2020/10/19 原:4
    Dim dblTot_O(2 To 6) As Double '其他
    Dim dblSum_Pre(2 To 6) As Double '上個月
    'end 2020/10/19
    Dim bolFormula As Boolean '是否設公式
    'Add by Amy 2017/07/18
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    Dim dblR As Double 'Add by Amy 2020/10/19 當月收文/結餘點數
    
    ReDim strRowN1(5)
    strRowN1 = Array("-", "目標點數", "達成點數", "達成率％", "保留點數", "實際點數")
    
    'Modify by Amy 2020/10/19  原:上月/本月「未報」改上月/本月「保留 」,實際達成改當月實績,加 當月收文/當月結餘
    ReDim strColN(7)
    strColN = Array("區別", "達成點數", "上月保留", "本月保留", "當月實績", "當月收文", "當月結餘")
    
    ReDim intWidth(UBound(strColN))
    intWidth = Array(13, 13, 13, 13, 13, 13, 13)
    'end 2020/10/19
    
    'Add by Amy 2015/06/09 +列印設定
    With wksrpt.PageSetup
        .Orientation = xlPortrait '直印
        .Zoom = False '縮放比例要設false,FitToPagesWide和FitToPagesTall才有作用
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With

    With wksrpt
        .Range(Chr(intField) & intCounter).Value = "智權部工作報告（" & Val(Right(txtCloseDate, 2)) & "月份）"
        .Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strColN)) & intCounter).HorizontalAlignment = xlCenter
        .Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strColN)) & intCounter).VerticalAlignment = xlBottom
        .Range(Chr(intField) & intCounter & ":" & Chr(intField + UBound(strColN)) & intCounter).MergeCells = True
        intCounter = intCounter + 1
        .Range(Chr(intField + UBound(strColN)) & intCounter).Value = ChangeWStringToTDateString(strSrvDate(1))
        intCounter = intCounter + 1
        
        TitleRow = intCounter
        '業務部達成情形
        .Range(Chr(intField) & intCounter).Value = "壹.業務部達成情形"
        intCounter = intCounter + 1
        
        '列
        For ii = 0 To UBound(strRowN1)
            If strRowN1(ii) = "-" Then
                .Range(Chr(intField) & intCounter + ii).Select
                xlsCustPoint.Selection.Borders(xlDiagonalDown).LineStyle = xlContinuous
            Else
                .Range(Chr(intField) & intCounter + ii).Value = strRowN1(ii)
            End If
            .Range(Chr(intField) & intCounter + ii).HorizontalAlignment = xlCenter
        Next ii
        intCounter = intCounter + UBound(strRowN1) + 2
        
        '=== 貳.各區達成情形 ====
        .Range(Chr(intField) & intCounter).Value = "貳.各區達成情形"
        intCounter = intCounter + 1
        
        '欄
        For ii = 0 To UBound(strColN)
            .Columns(Chr(intField + ii)).ColumnWidth = intWidth(ii)
            .Range(Chr(intField + ii) & intCounter).Value = strColN(ii)
            .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
        Next ii
        intCounter = intCounter + 1: StartRow = intCounter
        
        '資料(避免資料值小數三位,但顯示小數二位做加總導致與計算不同,故先四捨五入)
        'Modify by Amy 2020/10/19 上月/本明未報改上月/本明保留,實際達成改顯示當月實績,+當月收文/結餘
        Do While adoMon.EOF = False
            dblR = PUB_CountCP18(0, Val(txtCloseDate & "01") + 19110000, Val(txtCloseDate & "31") + 19110000, adoMon.Fields("ST15"), adoMon.Fields("ST15"))
            If Left(adoMon.Fields("ST15"), 1) = "S" And adoMon.Fields("ST15") <> "S00" Then
                .Range(Chr(intField + GetValueCN("區別")) & intCounter).Value = adoMon.Fields("DepN")
                .Range(Chr(intField + GetValueCN("區別")) & intCounter).HorizontalAlignment = xlCenter
                .Range(Chr(intField + GetValueCN("達成點數")) & intCounter).Value = "=" & Chr(intField + GetValueCN("上月保留")) & intCounter & _
                                                                                                                                       "+" & Chr(intField + GetValueCN("當月實績")) & intCounter & _
                                                                                                                                       "+" & Chr(intField + GetValueCN("當月結餘")) & intCounter & _
                                                                                                                                       "-" & Chr(intField + GetValueCN("本月保留")) & intCounter

                .Range(Chr(intField + GetValueCN("上月保留")) & intCounter).Value = Round(Val(adoMon.Fields("PreKeep")), 2)
                .Range(Chr(intField + GetValueCN("本月保留")) & intCounter).Value = Round(Val(adoMon.Fields("NowKeep")), 2)
                .Range(Chr(intField + GetValueCN("當月實績")) & intCounter).Value = Round(Val(adoMon.Fields("NowA")), 2)
                .Range(Chr(intField + GetValueCN("當月收文")) & intCounter).Value = Round(dblR, 2)
                .Range(Chr(intField + GetValueCN("當月結餘")) & intCounter).Value = Round(Val(adoMon.Fields("NowS")), 2)
                              
                intCounter = intCounter + 1
            Else
                'Mark by Amy 2021/05/27 11004月,因其他包含L的資料,故將其他不顯示-簡協理
'                '其他
'                dblTot_O(GetValueCN("上月保留")) = dblTot_O(GetValueCN("上月保留")) + Round(Val(adoMon.Fields("PreKeep")), 2)
'                dblTot_O(GetValueCN("本月保留")) = dblTot_O(GetValueCN("本月保留")) + Round(Val(adoMon.Fields("NowKeep")), 2)
'                dblTot_O(GetValueCN("當月實績")) = dblTot_O(GetValueCN("當月實績")) + Round(Val(adoMon.Fields("NowA")), 2)
'                dblTot_O(GetValueCN("當月收文")) = dblTot_O(GetValueCN("當月收文")) + Round(dblR, 2)
'                dblTot_O(GetValueCN("當月結餘")) = dblTot_O(GetValueCN("當月結餘")) + Round(Val(adoMon.Fields("NowS")), 2)
                
            End If
            '本月較上個月(for 壹.表資料)
            dblSum_Pre(GetValueCN("上月保留")) = dblSum_Pre(GetValueCN("上月保留")) + Round(Val(adoMon.Fields("PreM1")), 2)
            dblSum_Pre(GetValueCN("本月保留")) = dblSum_Pre(GetValueCN("本月保留")) + Round(Val(adoMon.Fields("PreM2")), 2)
            dblSum_Pre(GetValueCN("當月實績")) = dblSum_Pre(GetValueCN("當月實績")) + Round(Val(adoMon.Fields("PreM3")), 2)
            dblSum_Pre(GetValueCN("當月收文")) = dblSum_Pre(GetValueCN("當月收文")) + Round(dblR, 2)
            dblSum_Pre(GetValueCN("當月結餘")) = dblSum_Pre(GetValueCN("當月結餘")) + Round(Val(adoMon.Fields("PreM4")), 2)
            
            adoMon.MoveNext
        Loop
        'Mark by Amy 2021/05/27 11004月,因其他包含L的資料,故將其他不顯示-簡協理
'        '其他(非智權部的資料)
'        .Range(Chr(intField + GetValueCN("區別")) & intCounter).Value = "其他"
'        .Range(Chr(intField + GetValueCN("區別")) & intCounter).HorizontalAlignment = xlCenter
'        .Range(Chr(intField + GetValueCN("達成點數")) & intCounter).Value = "=" & Chr(intField + GetValueCN("上月保留")) & intCounter & _
'                                                                                                                               "+" & Chr(intField + GetValueCN("當月實績")) & intCounter & _
'                                                                                                                                "+" & Chr(intField + GetValueCN("當月結餘")) & intCounter & _
'                                                                                                                                 "-" & Chr(intField + GetValueCN("本月保留")) & intCounter
'
'        .Range(Chr(intField + GetValueCN("上月保留")) & intCounter).Value = dblTot_O(GetValueCN("上月保留"))
'        .Range(Chr(intField + GetValueCN("本月保留")) & intCounter).Value = dblTot_O(GetValueCN("本月保留"))
'        .Range(Chr(intField + GetValueCN("當月實績")) & intCounter).Value = dblTot_O(GetValueCN("當月實績"))
'        .Range(Chr(intField + GetValueCN("當月收文")) & intCounter).Value = "XXXXX" '先不抓-秀玲
'        .Range(Chr(intField + GetValueCN("當月收文")) & intCounter).HorizontalAlignment = xlCenter
'        .Range(Chr(intField + GetValueCN("當月結餘")) & intCounter).Value = dblTot_O(GetValueCN("當月結餘"))
        'end 2021/05/27
        'end 2020/10/19
         intCounter = intCounter + 1
        '合計
        .Range(Chr(intField + GetValueCN("區別")) & intCounter).Value = "合計"
        .Range(Chr(intField + GetValueCN("區別")) & intCounter).HorizontalAlignment = xlCenter
        For ii = 1 To UBound(strColN)
            .Range(Chr(intField + ii) & intCounter).Formula = "=Sum(" & Chr(intField + ii) & StartRow & ":" & Chr(intField + ii) & intCounter - 1 & ")"
        Next ii
        EndRow = intCounter
        
        '畫框
        .Range(Chr(intField) & StartRow - 1 & ":" & Chr(intField + UBound(strColN)) & intCounter).Select
        .Range(Chr(intField) & StartRow - 1 & ":" & Chr(intField + UBound(strColN)) & intCounter).NumberFormatLocal = "#,##0.00_ "
        xlsCustPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
        xlsCustPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
        xlsCustPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        xlsCustPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
        xlsCustPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
        xlsCustPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
        '=== 貳.各區達成情形 End ====
        
        '=== 壹.資料 ===
        intCounter = TitleRow + 1: intTp = 1
        
        '抓取目標點數
        If Val(Right(txtCloseDate, 2)) = 1 Then
            If Len(txtCloseDate) = 5 Then
                strMon = Val(Left(txtCloseDate, 3)) - 1
            Else
                strMon = Val(Left(txtCloseDate, 2)) - 1
            End If
            strMon = Val(strMon & "12") + 191100
        Else
            strMon = Val(txtCloseDate) + 191099
        End If
        'Modify by Amy 2021/05/27 原:SubStr(st15,1,1)<>'F' 11004月,因其他包含L的資料,故將其他不顯示(只顯示智權部)-簡協理
        strMon = "Select PE03,Sum(PE04) PE04 From Staff,PerFormance Where SubStr(st15,1,1)='S' And PE01(+)=ST01 And PE02(+)='TOT' " & _
                    "And PE03>=" & strMon & " And PE03<= " & Val(txtCloseDate) + 191100 & " Group by PE03 "
    
        If adoMon.State = adStateOpen Then adoMon.Close
        adoMon.CursorLocation = adUseClient
        adoMon.Open strMon, adoTaie, adOpenStatic, adLockReadOnly
        If adoMon.RecordCount > 0 Then
            Do While adoMon.EOF = False
                For ii = 0 To UBound(strRowN1)
                    bolFormula = False
                    Select Case ii
                        Case GetValueRN1("-")
                            strTp = ChgTxt(Val(Right(adoMon.Fields("PE03"), 2)))
                        Case GetValueRN1("目標點數")
                             strTp = adoMon.Fields("PE04")
                        Case GetValueRN1("達成點數")
                            If Val(Right(txtCloseDate, 2)) = Val(Right(adoMon.Fields("PE03"), 2)) Then
                                strTp = Chr(intField + GetValueCN("達成點數")) & EndRow
                                bolFormula = True
                            Else
                                'Modify by Amy 2020/10/19 實際達成 顯示成 當月實績,實際達成=當月實績+當月結餘
                                'strTp = dblSum_Pre(GetValueCN("上月保留")) + dblSum_Pre(GetValueCN("實際達成")) - dblSum_Pre(GetValueCN("本月保留"))
                                strTp = dblSum_Pre(GetValueCN("上月保留")) + dblSum_Pre(GetValueCN("當月實績")) + dblSum_Pre(GetValueCN("當月結餘")) - dblSum_Pre(GetValueCN("本月保留"))
                            End If
                        Case GetValueRN1("達成率％")
                             strTp = Chr(intField + intTp) & intCounter + GetValueRN1("達成點數") & "/" & Chr(intField + intTp) & intCounter + GetValueRN1("目標點數")
                            bolFormula = True
                        Case GetValueRN1("保留點數")
                            If intTp = 1 Then
                                strTp = dblSum_Pre(GetValueCN("本月保留"))
                            Else
                                strTp = Chr(intField + GetValueCN("本月保留")) & EndRow
                            End If
                            bolFormula = True
                        Case GetValueRN1("實際點數")
                            'Modify by Amy 2020/10/19 實際達成, 顯示成 當月實績,實際達成=當月實績+當月結餘
                            If intTp = 1 Then
                                strTp = dblSum_Pre(GetValueCN("當月實績")) + dblSum_Pre(GetValueCN("當月結餘"))
                            Else
                                strTp = Chr(intField + GetValueCN("當月實績")) & EndRow & "+" & Chr(intField + GetValueCN("當月結餘")) & EndRow
                                bolFormula = True
                            End If
                            'end 2020/10/19
                        Case Else
                    End Select
                    
                    If bolFormula = True Then
                        .Range(Chr(intField + intTp) & intCounter + ii).Formula = "=" & strTp
                    Else
                        .Range(Chr(intField + intTp) & intCounter + ii).Value = strTp
                    End If
                    If ii = GetValueRN1("-") Then
                        .Range(Chr(intField + intTp) & intCounter + ii).HorizontalAlignment = xlCenter
                    ElseIf ii = GetValueRN1("達成率％") Then
                        .Range(Chr(intField + intTp) & intCounter + ii).NumberFormatLocal = "0.00%"
                    Else
                        .Range(Chr(intField + intTp) & intCounter + ii).NumberFormatLocal = "#,##0.00_ "
                    End If
                Next ii
                intTp = intTp + 1
                adoMon.MoveNext
            Loop
            '本月較上月
             For ii = 0 To UBound(strRowN1)
                If ii = GetValueRN1("-") Then
                    .Range(Chr(intField + intTp) & intCounter + ii).Value = "本月較上月"
                ElseIf ii <> GetValueRN1("目標點數") Then
                    .Range(Chr(intField + intTp) & intCounter + ii).Formula = "=" & Chr(intField + 2) & intCounter + ii & "-" & Chr(intField + 1) & intCounter + ii
                End If
                .Range(Chr(intField + intTp) & intCounter + ii & ":" & Chr(intField + intTp + 1) & intCounter + ii).MergeCells = True
                If ii = GetValueRN1("-") Then
                    .Range(Chr(intField + intTp) & intCounter + ii).HorizontalAlignment = xlCenter
                Else
                    .Range(Chr(intField + intTp) & intCounter + ii & ":" & Chr(intField + intTp + 1) & intCounter + ii).HorizontalAlignment = xlRight
                End If
                
                If ii = GetValueRN1("目標點數") Then
                    '畫\要在合併儲存格之後
                    .Range(Chr(intField + intTp) & intCounter + ii & ":" & Chr(intField + intTp + 1) & intCounter + ii).Select
                    xlsCustPoint.Selection.Borders(xlDiagonalDown).LineStyle = xlContinuous
                End If
            Next ii
            '畫框
            .Range(Chr(intField) & TitleRow + 1 & ":" & Chr(intField + 4) & TitleRow + 1 + UBound(strRowN1)).Select
            xlsCustPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            '=== 壹.資料 End ===
            
            'Add by Amy 2017/07/18 當有若有結餘點數加備註
            intCounter = EndRow + 2: strTp = "": strExc(0) = ""
            'Modify by Amy 2020/10/19 原:GetPoint(1…),程式都改抓GetPoint故調整,txtCloseDate改抓變數
            'Memo by Amy 2021/05/27 改basQuery 11004月,因其他包含L的資料,故將其他不顯示(只顯示智權部)-簡協理
            strQ = GetPoint(1.41, Val(StrDate2), Val(StrDate2), , , , , Me.Name, True)
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 1 Then
                intCounter = intCounter + 1
                If Len(StrDate2) = 5 Then
                    strTp = Left(StrDate2, 3) & "/" & Mid(StrDate2, 4)
                Else
                    strTp = Left(StrDate2, 2) & "/" & Mid(StrDate2, 3)
                End If
            'end 2020/10/19
                strTp = "註：" & strTp
                wksrpt.Range(Chr(intField) & intCounter).Value = strTp & "月結餘點數 " & Format(RsQ.Fields(0), FDollar) & " 點"
                wksrpt.Range(Chr(intField) & intCounter).Font.Color = RGB(255, 0, 0)
            End If
            'Modify by Amy 2020/10/19 原:GetPoint(1…),程式都改抓GetPoint故調整
            'Memo by Amy 2021/05/27 改basQuery 11004月,因其他包含L的資料,故將其他不顯示(只顯示智權部)-簡協理
            strQ = GetPoint(1.41, Val(strDate1), Val(strDate1), , , , , Me.Name, True)
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 1 Then
                If strTp = MsgText(601) Then
                    strTp = "註："
                Else
                    strTp = "　　"
                End If
                If Len(strDate1) = 5 Then
                    strTp = strTp & Left(strDate1, 3) & "/" & Mid(strDate1, 4)
                Else
                    strTp = strTp & Left(strDate1, 2) & "/" & Mid(strDate1, 3)
                End If
                intCounter = intCounter + 1
                wksrpt.Range(Chr(intField) & intCounter).Value = strTp & "月結餘點數 " & Format(RsQ.Fields(0), FDollar) & " 點"
                wksrpt.Range(Chr(intField) & intCounter).Font.Color = RGB(255, 0, 0)
            End If
            'end 2017/07/18
            '改 Sheet名稱
            .Name = "工作報告"
        End If
        
    End With
End Sub

Private Function doQuery2() As Boolean
    Dim strSP(3) As String, strST(3) As String, strSCFP(1) As String, strSCFT(3) As String
    Dim strGP(3) As String, strGT(3) As String, strGCFP(1) As String, strGCFT(3) As String
    Dim strWhere(1) As String
    
On Error GoTo ErrHnd
    intCounter = 1
    
    If bolOpenXls = False Then fn_CreateExcel
    'Modify by Amy 2018/04/16 for 工作表名稱改為中文
    If strWkName = MsgText(601) Then strWkName = Left(xlsCustPoint.Worksheets(1).Name, Len(xlsCustPoint.Worksheets(1).Name) - 1)
    Set wksrpt = xlsCustPoint.Worksheets(strWkName & intXlsSheet)
    'end 2018/04/16
    wksrpt.Activate
    
    If Val(Right(txtCloseDate, 2)) = 1 Then
        If Len(txtCloseDate) = 5 Then
            strWhere(0) = " And cp27 >= " & Val(Left(txtCloseDate, 3)) + 1910 & "1201 And cp27 <= " & Val(txtCloseDate) + 191100 & "31" '發文
            strWhere(1) = " And cp05 >= " & Val(Left(txtCloseDate, 3)) + 1910 & "1201 And cp05 <= " & Val(txtCloseDate) + 191100 & "31" '收文
        Else
            strWhere(0) = " And cp27 >= " & Val(Left(txtCloseDate, 2)) + 1910 & "1201 And cp27 <= " & Val(txtCloseDate) + 191100 & "31" '發文
            strWhere(1) = " And cp05 >= " & Val(Left(txtCloseDate, 2)) + 1910 & "1201 And cp05 <= " & Val(txtCloseDate) + 191100 & "31" '收文
        End If
    Else
        strWhere(0) = " And cp27 >= " & Val(txtCloseDate) + 191099 & "01 And cp27 <= " & Val(txtCloseDate) + 191100 & "31" '發文
        strWhere(1) = " And cp05 >= " & Val(txtCloseDate) + 191099 & "01 And cp05 <= " & Val(txtCloseDate) + 191100 & "31" '收文
    End If
    
    '*** P ***
    '發文
    strMon = "Select Distinct SubStr(cp27,1,6) stDate From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' " & strWhere(0)
                
    strSP(0) = "Select SubStr(cp27,1,6) V10,Count(*) V11 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And InStr('" & NewCasePtyList & "',cp10)>0 " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(1) = "Select SubStr(cp27,1,6) V20,Count(*) V21 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='107' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(2) = "Select SubStr(cp27,1,6) V30,Count(*) V31 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='500' And cp10<='599' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(3) = "Select SubStr(cp27,1,6) V40,Count(*) V41 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='800' And cp10<='899' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
                            
    '收文
    strGP(0) = "Select SubStr(cp05,1,6) V50,Count(*) V51 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And InStr('" & NewCasePtyList & "',cp10)>0 " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strGP(1) = "Select SubStr(cp05,1,6) V60,Count(*) V61 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                 "And SubStr(cp12,1,1)='S' And cp10='107' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
    
    strGP(2) = "Select SubStr(cp05,1,6) V70,Count(*) V71 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='500' And cp10<='599' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
    
    strGP(3) = "Select SubStr(cp05,1,6) V80,Count(*) V81 From CaseProgress Where cp01='P' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='800' And cp10<='899' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strMon = "Select stDate,Nvl(V11,0) C10,Nvl(V21,0) C20,Nvl(V31,0) C30,Nvl(V41,0) C40,Nvl(V51,0) C11,Nvl(V61,0) C21,Nvl(V71,0) C31,Nvl(V81,0) C41 " & _
                "From (" & strMon & "),(" & strSP(0) & " ),(" & strSP(1) & "),(" & strSP(2) & "),(" & strSP(3) & "),(" & strGP(0) & " ),(" & strGP(1) & "),(" & strGP(2) & "),(" & strGP(3) & ") " & _
                "Where stDate=V10(+) And stDate=V20(+) And stDate=V30(+) And stDate=V40(+) And stDate=V50(+) And stDate=V60(+) And stDate=V70(+) And stDate=V80(+)"
            
    If adoMon.State = adStateOpen Then adoMon.Close
    adoMon.CursorLocation = adUseClient
    adoMon.Open strMon, adoTaie, adOpenStatic, adLockReadOnly
    SaveExcel2 "P", adoMon.RecordCount
   
   '*** T ***
     strMon = "Select Distinct SubStr(cp27,1,6) stDate From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' " & strWhere(0)
                
    strSP(0) = "Select SubStr(cp27,1,6) V10,Count(*) V11 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='101' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(1) = "Select SubStr(cp27,1,6) V20,Count(*) V21 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='102' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(2) = "Select SubStr(cp27,1,6) V30,Count(*) V31 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='400' And cp10<='499' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(3) = "Select SubStr(cp27,1,6) V40,Count(*) V41 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='600' And cp10<='699' And cp10 Not In('611','612','613','614','615') " & strWhere(0) & " Group by SubStr(cp27,1,6)"
                            
    '收文
    strGP(0) = "Select SubStr(cp05,1,6) V50,Count(*) V51 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='101' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strGP(1) = "Select SubStr(cp05,1,6) V60,Count(*) V61 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                 "And SubStr(cp12,1,1)='S' And cp10='102' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
    
    strGP(2) = "Select SubStr(cp05,1,6) V70,Count(*) V71 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='400' And cp10<='499' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
    
    strGP(3) = "Select SubStr(cp05,1,6) V80,Count(*) V81 From CaseProgress Where cp01='T' And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10>='600' And cp10<='699' And cp10 Not In('611','612','613','614','615') " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strMon = "Select stDate,Nvl(V11,0) C10,Nvl(V21,0) C20,Nvl(V31,0) C30,Nvl(V41,0) C40,Nvl(V51,0) C11,Nvl(V61,0) C21,Nvl(V71,0) C31,Nvl(V81,0) C41 " & _
                "From (" & strMon & "),(" & strSP(0) & " ),(" & strSP(1) & "),(" & strSP(2) & "),(" & strSP(3) & "),(" & strGP(0) & " ),(" & strGP(1) & "),(" & strGP(2) & "),(" & strGP(3) & ") " & _
                "Where stDate=V10(+) And stDate=V20(+) And stDate=V30(+) And stDate=V40(+) And stDate=V50(+) And stDate=V60(+) And stDate=V70(+) And stDate=V80(+)"
            
    If adoMon.State = adStateOpen Then adoMon.Close
    adoMon.CursorLocation = adUseClient
    adoMon.Open strMon, adoTaie, adOpenStatic, adLockReadOnly
    SaveExcel2 "T", adoMon.RecordCount
   
     '*** CFP ***
     strMon = "Select Distinct SubStr(cp27,1,6) stDate From CaseProgress Where cp01='CFP'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' " & strWhere(0)
                
    strSP(0) = "Select SubStr(cp27,1,6) V10,Count(*) V11 From CaseProgress Where cp01='CFP'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And InStr('" & NewCasePtyList & "',cp10)>0 " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(1) = "Select SubStr(cp27,1,6) V20,Count(*) V21 From CaseProgress Where cp01='CFP'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='107' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
                           
    '收文
    strGP(0) = "Select SubStr(cp05,1,6) V50,Count(*) V51 From CaseProgress Where cp01='CFP'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And InStr('" & NewCasePtyList & "',cp10)>0 " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strGP(1) = "Select SubStr(cp05,1,6) V60,Count(*) V61 From CaseProgress Where cp01='CFP'And cp09<'B' And Nvl(cp57,0)=0 " & _
                 "And SubStr(cp12,1,1)='S' And cp10='107' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strMon = "Select stDate,Nvl(V11,0) C10,Nvl(V21,0) C20,Nvl(V51,0) C11,Nvl(V61,0) C21 " & _
                "From (" & strMon & "),(" & strSP(0) & " ),(" & strSP(1) & "),(" & strGP(0) & " ),(" & strGP(1) & ") " & _
                "Where stDate=V10(+) And stDate=V20(+) And stDate=V50(+) And stDate=V60(+)"
            
    If adoMon.State = adStateOpen Then adoMon.Close
    adoMon.CursorLocation = adUseClient
    adoMon.Open strMon, adoTaie, adOpenStatic, adLockReadOnly
    SaveExcel2 "CFP", adoMon.RecordCount
    
     '*** CFT ***
     strMon = "Select Distinct SubStr(cp27,1,6) stDate From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' " & strWhere(0)
                
    strSP(0) = "Select SubStr(cp27,1,6) V10,Count(*) V11 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='101' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(1) = "Select SubStr(cp27,1,6) V20,Count(*) V21 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='202' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(2) = "Select SubStr(cp27,1,6) V30,Count(*) V31 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='105' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
    
    strSP(3) = "Select SubStr(cp27,1,6) V40,Count(*) V41 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='701' " & strWhere(0) & " Group by SubStr(cp27,1,6)"
                            
    '收文
    strGP(0) = "Select SubStr(cp05,1,6) V50,Count(*) V51 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='101' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strGP(1) = "Select SubStr(cp05,1,6) V60,Count(*) V61 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                 "And SubStr(cp12,1,1)='S' And cp10='202' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
    
    strGP(2) = "Select SubStr(cp05,1,6) V70,Count(*) V71 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='105' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
    
    strGP(3) = "Select SubStr(cp05,1,6) V80,Count(*) V81 From CaseProgress Where cp01='CFT'And cp09<'B' And Nvl(cp57,0)=0 " & _
                "And SubStr(cp12,1,1)='S' And cp10='701' " & strWhere(1) & " Group by SubStr(cp05,1,6)"
                
    strMon = "Select stDate,Nvl(V11,0) C10,Nvl(V21,0) C20,Nvl(V31,0) C30,Nvl(V41,0) C40,Nvl(V51,0) C11,Nvl(V61,0) C21,Nvl(V71,0) C31,Nvl(V81,0) C41 " & _
                "From (" & strMon & "),(" & strSP(0) & " ),(" & strSP(1) & "),(" & strSP(2) & "),(" & strSP(3) & "),(" & strGP(0) & " ),(" & strGP(1) & "),(" & strGP(2) & "),(" & strGP(3) & ") " & _
                "Where stDate=V10(+) And stDate=V20(+) And stDate=V30(+) And stDate=V40(+) And stDate=V50(+) And stDate=V60(+) And stDate=V70(+) And stDate=V80(+)"
            
    If adoMon.State = adStateOpen Then adoMon.Close
    adoMon.CursorLocation = adUseClient
    adoMon.Open strMon, adoTaie, adOpenStatic, adLockReadOnly
    SaveExcel2 "CFT", adoMon.RecordCount
    
    intCounter = intCounter + 1
    wksrpt.Range(Chr(intField) & intCounter).Value = "*取消收文案件不計入"
    wksrpt.Range(Chr(intField) & intCounter).Font.Color = RGB(255, 0, 0)
    
    wksrpt.Name = "收.發文量比較表"
    doQuery2 = True
    adoMon.Close
    Exit Function
    
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

Private Sub SaveExcel2(ByVal SysKind As String, ByVal intRCount As Integer)
    Dim strTp As String, strTp2 As String
    Dim StartRow As Integer '新的系統別表格/開始列 'bolIsFirst As Boolean,
    Dim bolSetWidth As Boolean '是否已設寬
   
    Select Case SysKind
        Case "P"
            ReDim strColN(7)
            strColN = Array("案件性質", "P", "申請案", "再審", "救濟案", "爭議案", "合計")
        Case "T"
            ReDim strColN(7)
            strColN = Array("案件性質", "T", "申請案", "延展", "救濟案", "爭議案", "合計")
        Case "CFP"
            ReDim strColN(5)
            strColN = Array("案件性質", "CFP", "申請案", "答辯", "合計")
            ReDim intWidth(UBound(strColN))
            intWidth = Array(13, 13, 13, 13, 13)
            bolSetWidth = True
        Case "CFT"
            ReDim strColN(7)
            strColN = Array("案件性質", "CFT", "申請案", "答辯", "使用宣誓", "領證", "合計")
        Case Else
    End Select
    
    If bolSetWidth = False Then
        ReDim intWidth(UBound(strColN))
        intWidth = Array(11.5, 11.5, 11.5, 11.5, 11.5, 11.5, 11.5)
    End If
    
    'Add by Amy 2015/06/09 +列印設定
    With wksrpt.PageSetup
        .Orientation = xlPortrait '直印
        .Zoom = False '縮放比例要設false,FitToPagesWide和FitToPagesTall才有作用
        .FitToPagesWide = 1
        .FitToPagesTall = 1
    End With
    
    With wksrpt
        If SysKind = "P" Then
            .Range(Chr(intField) & intCounter).Value = "案件收.發文量比較"
            intCounter = intCounter + 1: TitleRow = intCounter
        End If
        
        For ii = 0 To UBound(strColN)
            If ii = 1 Then .Range(Chr(intField + ii) & intCounter).Font.Bold = True
            .Columns(Chr(intField + ii)).ColumnWidth = intWidth(ii)
            .Range(Chr(intField + ii) & intCounter).Value = strColN(ii)
            .Range(Chr(intField + ii) & intCounter).HorizontalAlignment = xlCenter
        Next ii
        intCounter = intCounter + 1: StartRow = intCounter
        
        If intRCount > 0 Then
            Do While adoMon.EOF = False
                For ii = 0 To UBound(strColN)
                    Select Case ii
                        Case GetValueCN("案件性質")
                            If intCounter = StartRow Then
                                strTp = "發文量"
                                strTp2 = "收文量"
                            Else
                                strTp = "": strTp2 = ""
                            End If
                        Case GetValueCN(SysKind)
                            strTp = ChgTxt(Val(Right(adoMon.Fields("StDate"), 2)))
                            strTp2 = strTp
                        Case GetValueCN("合計")
                            strTp = Chr(intField + GetValueCN("申請案")) & intCounter & ":" & Chr(intField + UBound(strColN) - 1) & intCounter
                            strTp2 = Chr(intField + GetValueCN("申請案")) & intCounter + 2 & ":" & Chr(intField + UBound(strColN) - 1) & intCounter + 2
                        Case Else
                            strTp = adoMon.Fields(ii - 1)
                            If SysKind = "CFP" Then
                                strTp2 = adoMon.Fields(ii + 1)
                            Else
                                strTp2 = adoMon.Fields(ii + 3)
                            End If
                    End Select
                    
                    If ii = GetValueCN("合計") Then
                         .Range(Chr(intField + ii) & intCounter).Formula = "=Sum(" & strTp & ")"
                         .Range(Chr(intField + ii) & intCounter + 2).Formula = "=Sum(" & strTp2 & ")"
                    Else
                        .Range(Chr(intField + ii) & intCounter).Value = strTp
                        .Range(Chr(intField + ii) & intCounter + 2).Value = strTp2
                    End If
                Next ii
                intCounter = intCounter + 1
                adoMon.MoveNext
            Loop
            intCounter = intCounter + 1
            '畫框
            .Range(Chr(intField) & StartRow - 1 & ":" & Chr(intField + UBound(strColN)) & intCounter).HorizontalAlignment = xlCenter
            .Range(Chr(intField) & StartRow - 1 & ":" & Chr(intField + UBound(strColN)) & intCounter).Select
            xlsCustPoint.Selection.Borders(xlEdgeLeft).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlEdgeTop).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlEdgeRight).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlInsideVertical).LineStyle = xlContinuous
            xlsCustPoint.Selection.Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Range(Chr(intField) & StartRow + adoMon.RecordCount & ":" & Chr(intField + UBound(strColN)) & StartRow + adoMon.RecordCount).Select
            xlsCustPoint.Selection.Borders(xlEdgeTop).Weight = xlMedium
            intCounter = intCounter + 2
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm210150 = Nothing
   Set adoMon = Nothing
   Set adoPMon = Nothing
End Sub

Private Sub txtCloseDate_GotFocus()
    TextInverse txtCloseDate
    CloseIme
End Sub

Private Sub txtCloseDate_Validate(Cancel As Boolean)
     If txtCloseDate <> MsgText(601) Then
        If ChkDate(txtCloseDate & "01") = False Then
            txtCloseDate_GotFocus
            Cancel = True
        End If
    End If
End Sub

Private Function GetValueCN(pColN As String) As Integer
    Dim r As Integer
    For r = 0 To UBound(strColN)
       If UCase(strColN(r)) = UCase(pColN) Then
          GetValueCN = r
          Exit For
       End If
    Next r
End Function

Private Function GetValueRN1(pRowN As String) As Integer
    Dim r As Integer
    For r = 1 To UBound(strRowN1)
       If UCase(strRowN1(r)) = UCase(pRowN) Then
          GetValueRN1 = r
          Exit For
       End If
    Next r
End Function

'數字月份轉成國字
Private Function ChgTxt(intValue As Integer) As String
    Select Case intValue
        Case 1
            ChgTxt = "一"
        Case 2
            ChgTxt = "二"
        Case 3
            ChgTxt = "三"
        Case 4
            ChgTxt = "四"
        Case 5
            ChgTxt = "五"
        Case 6
            ChgTxt = "六"
        Case 7
            ChgTxt = "七"
        Case 8
            ChgTxt = "八"
        Case 9
            ChgTxt = "九"
        Case 10
            ChgTxt = "十"
        Case 11
            ChgTxt = "十一"
        Case 12
            ChgTxt = "十二"
       Case Else
    End Select
    If ChgTxt <> MsgText(601) Then
        If intValue > 9 Then
            ChgTxt = ChgTxt & "月"
        Else
            ChgTxt = ChgTxt & "　月"
        End If
    End If
End Function

