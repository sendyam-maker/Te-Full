VERSION 5.00
Begin VB.Form frm090201_3_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "工作進度資料維護(專利)"
   ClientHeight    =   3200
   ClientLeft      =   2040
   ClientTop       =   1920
   ClientWidth     =   3690
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3200
   ScaleWidth      =   3690
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1800
      TabIndex        =   1
      Top             =   24
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   2592
      TabIndex        =   0
      Top             =   24
      Width           =   1092
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "lbl1(8)"
      Height          =   180
      Index           =   8
      Left            =   1800
      TabIndex        =   20
      Top             =   990
      Width           =   1440
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "待會修件數：                                               件"
      Height          =   180
      Left            =   84
      TabIndex        =   19
      Top             =   990
      Width           =   3570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "括弧內為舊制算法"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   150
      TabIndex        =   18
      Top             =   2745
      Width           =   1440
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "當月已完稿設計案件：                               件"
      Height          =   180
      Left            =   84
      TabIndex        =   17
      Top             =   1500
      Width           =   3570
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label2"
      Height          =   180
      Index           =   3
      Left            =   2070
      TabIndex        =   16
      Top             =   1485
      Width           =   1155
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "可辦非設計案件：                                       件"
      Height          =   180
      Left            =   84
      TabIndex        =   15
      Top             =   504
      Width           =   3564
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label4"
      Height          =   180
      Index           =   0
      Left            =   1788
      TabIndex        =   14
      Top             =   504
      Width           =   1440
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "可辦設計案件：                                           件"
      Height          =   180
      Left            =   84
      TabIndex        =   13
      Top             =   756
      Width           =   3564
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label6"
      Height          =   180
      Index           =   1
      Left            =   1572
      TabIndex        =   12
      Top             =   744
      Width           =   1656
   End
   Begin VB.Label Label7 
      BackStyle       =   0  '透明
      Caption         =   "當月已完稿非設計案件：                           件"
      Height          =   180
      Left            =   84
      TabIndex        =   11
      Top             =   1245
      Width           =   3570
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label8"
      Height          =   180
      Index           =   2
      Left            =   2265
      TabIndex        =   10
      Top             =   1245
      Width           =   960
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "當日法定期限件數：                                   件"
      Height          =   180
      Left            =   84
      TabIndex        =   9
      Top             =   1755
      Width           =   3570
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label10"
      Height          =   180
      Index           =   4
      Left            =   1770
      TabIndex        =   8
      Top             =   1755
      Width           =   1455
   End
   Begin VB.Label Label11 
      BackStyle       =   0  '透明
      Caption         =   "當月發文件數：                                           件"
      Height          =   180
      Left            =   84
      TabIndex        =   7
      Top             =   2010
      Width           =   3570
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label12"
      Height          =   180
      Index           =   5
      Left            =   1440
      TabIndex        =   6
      Top             =   2010
      Width           =   1785
   End
   Begin VB.Label Label13 
      BackStyle       =   0  '透明
      Caption         =   "當月發文點數：                                           點"
      Height          =   180
      Left            =   84
      TabIndex        =   5
      Top             =   2250
      Width           =   3570
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label14"
      Height          =   180
      Index           =   6
      Left            =   1455
      TabIndex        =   4
      Top             =   2250
      Width           =   1770
   End
   Begin VB.Label Label15 
      BackStyle       =   0  '透明
      Caption         =   "超過承辦期限數：                                       件"
      Height          =   180
      Left            =   84
      TabIndex        =   3
      Top             =   2505
      Width           =   3570
   End
   Begin VB.Label lbl1 
      Alignment       =   1  '靠右對齊
      Caption         =   "Label16"
      Height          =   180
      Index           =   7
      Left            =   1590
      TabIndex        =   2
      Top             =   2490
      Width           =   1635
   End
End
Attribute VB_Name = "frm090201_3_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Lydia 2022/02/14 Form2.0已檢查 (無需修改的物件)
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 21) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String
Dim PLeft(0 To 21) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer
'Add By Cheng 2003/05/08
Public m_strYear As String, m_strMonth As String

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     PrintData
Case 1
     Me.Hide
     frm090201_2.Show
     Unload Me
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Process
End Sub


Sub Process()
'add by nickc 2005/05/04
Dim tmpLBL As String
Dim strDateS As String 'Add by Sindy 2013/9/30
Dim strDateE As String 'Add by Sindy 2013/9/30

'Modify By Cheng 2003/05/08
'strSQL = "SELECT SUM(DECODE(R111010,0,0,R111010)),SUM(DECODE(R111011,0,0,R111011)),SUM(DECODE(R111012,0,0,R111012)),SUM(DECODE(R111013,0,0,R111013)),SUM(DECODE(R111009,0,0,R111009)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),sum(decode(r111008,0,0,r111008)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Trim(frm090201_2.Combo1.Text) & "' AND R111002='2' "
'edit by nickc 2005/05/04
'strSQL = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),sum(r111008) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Trim(Left("" & frm090201_2.Combo1.Text, 6)) & "' AND R111002='2' "
strSql = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),sum(r111008),SUM(R111021),SUM(R111022),SUM(R111023),SUM(R111024),SUM(R111020),SUM(R111015),SUM(R111016),sum(r111019) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Trim(Left("" & frm090201_2.Combo1.Text, 6)) & "' AND R111002='2' "
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        For i = 0 To 7
            lbl1(i) = CheckStr(.Fields(i))
            If Len(Trim(lbl1(i))) = 0 Then
                lbl1(i) = "0"
            End If
        Next i
        'add by nickc 2005/05/04
        For i = 8 To 15
            tmpLBL = Format(CheckStr(.Fields(i)), "###,##0.00")
            If Len(Trim(tmpLBL)) = 0 Then
                tmpLBL = "0"
            End If
            'Modify by Morgan 2010/5/11 改舊制在內
            'lbl1(i - 8) = lbl1(i - 8) & "(" & tmpLBL & ")"
            lbl1(i - 8) = tmpLBL & "(" & lbl1(i - 8) & ")"
        Next i
'         'Add by Sindy 2013/9/30 當月會修次數
'         lbl1(8) = "0"
'         strDateS = Val(m_strYear) + 1911 & Format(m_strMonth, "00") & "01"
'         strDateE = Val(m_strYear) + 1911 & Format(m_strMonth, "00") & "31"
'         strSql = "select count(eep01) From empelectronprocess where eep05='" & Trim(Left("" & frm090201_2.Combo1.Text, 6)) & "'" & _
'                  " and eep04='" & EMP_會修 & "' and eep06 between " & strDateS & " and " & strDateE
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            lbl1(8) = RsTemp.Fields(0)
'         End If
'         '2013/9/30 END
         'Add by Sindy 2013/10/23 待會修件數
         lbl1(8) = "0"
         'Modify By Sindy 2016/9/5 and cp57 is null and cp27 is null ==> and cp158=0 and cp159=0
         strSql = "select e2.eep01,e2.eep02,e2.eep04,e2.eep05 from" & _
                  " (select eep01,max(eep02) eep02" & _
                  " from EmpElectronProcess,caseprogress where eep01=cp09(+) and cp158=0 and cp159=0 and eep04 in('" & EMP_送會 & "','" & EMP_會修 & "')" & _
                  " group by eep01) e1,EmpElectronProcess e2" & _
                  " where e1.eep01=e2.eep01(+) and e1.eep02=e2.eep02(+)" & _
                  " and e2.eep04='" & EMP_會修 & "'" & _
                  " and e2.eep05='" & Trim(Left("" & frm090201_2.Combo1.Text, 6)) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            lbl1(8) = RsTemp.RecordCount
         End If
         '2013/9/30 END
    End If
End With
CheckOC
End Sub

Sub PrintData()
PrintData2
End Sub

Sub GetPleft_2()
'定陣列
'字 SIZE = 9
'1 WORD = 180 PIX
'0.5 WORD = 90 PIX
'SPACE = 90 PIX
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0
PLeft(2) = PLeft(1) + (2.5 * 180)
PLeft(3) = PLeft(2) + (2.5 * 180)
PLeft(4) = PLeft(3) + (4.5 * 180)
PLeft(5) = PLeft(4) + (8 * 180)
PLeft(6) = PLeft(5) + (10.5 * 180)
PLeft(7) = PLeft(6) + (2 * 180)
PLeft(8) = PLeft(7) + (3.5 * 180)
PLeft(9) = PLeft(8) + (4.5 * 180)
PLeft(10) = PLeft(9) + (4.5 * 180)
PLeft(11) = PLeft(10) + (4.5 * 180)
PLeft(12) = PLeft(11) + (4.5 * 180)
PLeft(13) = PLeft(12) + (4.5 * 180)
PLeft(14) = PLeft(13) + (4.5 * 180)
PLeft(15) = PLeft(14) + (4.5 * 180)
PLeft(16) = PLeft(15) + (4.5 * 180)
PLeft(17) = PLeft(16) + (4.5 * 180)
PLeft(18) = PLeft(17) + (4.5 * 180)
PLeft(19) = PLeft(18) + (2.5 * 180)
PLeft(20) = PLeft(19) + (5.5 * 180)
End Sub

Sub ShowLine()
Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
iPrint = iPrint + 300
End Sub

Sub PrintData2()
strSql = "SELECT DISTINCT R110001 FROM R090614 WHERE ID='" & strUserNum & "' "
CheckOC2
Page = 1
adoRecordset1.CursorLocation = adUseClient
'Modify By Cheng 2003/05/08
'adoRecordset1.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
adoRecordset1.Open strSql, adoEng, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    adoRecordset1.MoveFirst
    Do While adoRecordset1.EOF = False
        strTemp3 = CheckStr(adoRecordset1.Fields(0))
        PrintData2_1 (CheckStr(adoRecordset1.Fields(0)))
        PrintEnd2_1 (CheckStr(adoRecordset1.Fields(0)))
        adoRecordset1.MoveNext
        If adoRecordset1.EOF = False Then
            Page = Page + 1
            Printer.NewPage
        End If
    Loop
End If
CheckOC2
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintData2_1(Strindex As String)
If Len(Strindex) = 0 Then
    strSql = "SELECT * FROM R090614 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='')  order by r110002 "
Else
    strSql = "SELECT * FROM R090614 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "'  order by r110002 "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle_2
        Do While .EOF = False
            For i = 0 To 20
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(5) = StrToStr(strTemp(5), 10)
            strTemp(7) = StrToStr(strTemp(7), 3)
            strTemp(8) = StrToStr(strTemp(8), 4)
            strTemp(15) = StrToStr(strTemp(15), 3)
            strTemp(19) = StrToStr(strTemp(19), 5)
            strTemp(20) = StrToStr(strTemp(20), 3)
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle_2
            End If
            PrintDatil_2
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintDatil_2() '列印資料

For i = 1 To 20
    If i = 1 Or i = 18 Then
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "####0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "####0")
    Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
    End If
Next i
iPrint = iPrint + 300
End Sub

Sub PrintEnd2_1(Strindex As String)
'列印結尾

ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
End If
If Len(Strindex) = 0 Then
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111010,0,0,R111010)),SUM(DECODE(R111011,0,0,R111011)),SUM(DECODE(R111012,0,0,R111012)),SUM(DECODE(R111013,0,0,R111013)),SUM(DECODE(R111009,0,0,R111009)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111008,0,0,R111008)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
    strSql = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008),SUM(R111021),SUM(R111022),SUM(R111023),SUM(R111024),SUM(R111020),SUM(R111015),SUM(R111016),SUM(R111019) FROM R090614_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') AND R111002='2' "
Else
    'Modify By Cheng 2003/05/08
'    strSQL = "SELECT SUM(DECODE(R111010,0,0,R111010)),SUM(DECODE(R111011,0,0,R111011)),SUM(DECODE(R111012,0,0,R111012)),SUM(DECODE(R111013,0,0,R111013)),SUM(DECODE(R111009,0,0,R111009)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)),SUM(DECODE(R111008,0,0,R111008)) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
    strSql = "SELECT SUM(R111010),SUM(R111011),SUM(R111012),SUM(R111013),SUM(R111009),SUM(R111004),SUM(R111005),SUM(R111008),SUM(R111021),SUM(R111022),SUM(R111023),SUM(R111024),SUM(R111020),SUM(R111015),SUM(R111016),SUM(R111019) FROM R090614_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' AND R111002='2' "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    'Modify By Cheng 2003/05/08
'    .Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
    .Open strSql, adoEng, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "可辦非設計案件：" & Format(IIf(CheckStr(.Fields(0)) = "", "0", CheckStr(.Fields(0))), "###,###,###,###,##0") & " 件"
        Printer.Print "可辦非設計案件：" & Format("0" & CheckStr(.Fields(0)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(8)), "###,###,###,###,##0.00") & ") 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "可辦設計案件：" & Format(IIf(CheckStr(.Fields(1)) = "", "0", CheckStr(.Fields(1))), "###,###,###,###,##0") & " 件"
        Printer.Print "可辦設計案件：" & Format("0" & CheckStr(.Fields(1)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(9)), "###,###,###,###,##0.00") & ") 件"
        
        'Add By Sindy 2013/9/30
        Printer.CurrentX = 14000
        Printer.CurrentY = iPrint
        Printer.Print "待會修件數：" & Format("0" & CheckStr(lbl1(8)), "###,###,###,###,##0") & " 件"
        '2013/9/30 END
        
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月已完稿非設計件數：" & Format(IIf(CheckStr(.Fields(2)) = "", "0", CheckStr(.Fields(2))), "###,###,###,###,##0") & " 件"
        Printer.Print "本月已完稿非設計件數：" & Format("0" & CheckStr(.Fields(2)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(10)), "###,###,###,###,##0.00") & ") 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月已完稿設計件數：" & Format(IIf(CheckStr(.Fields(3)) = "", "0", CheckStr(.Fields(3))), "###,###,###,###,##0") & " 件"
        Printer.Print "本月已完稿設計件數：" & Format("0" & CheckStr(.Fields(3)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(11)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "當日法定期限之件數：" & Format(IIf(CheckStr(.Fields(4)) = "", "0", CheckStr(.Fields(4))), "###,###,###,###,##0") & " 件"
        'Modified by Morgan 2015/9/30 改為本所期限(實際是統計所限過期的)
        Printer.Print "當日本所期限之件數：" & Format("0" & CheckStr(.Fields(4)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(12)), "###,###,###,###,##0.00") & ") 件"
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月發文件數：" & Format(IIf(CheckStr(.Fields(5)) = "", "0", CheckStr(.Fields(5))), "###,###,###,###,##0") & " 件 "
        Printer.Print "本月發文件數：" & Format("0" & CheckStr(.Fields(5)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(13)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        Printer.CurrentX = 7000
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "本月發文點數：" & Format(IIf(CheckStr(.Fields(6)) = "", "0", CheckStr(.Fields(6))), "###,###,###,###,##0.00") & " 點"
        Printer.Print "本月發文點數：" & Format("0" & CheckStr(.Fields(6)), "###,###,###,###,##0.00") & "(" & Format("0" & CheckStr(.Fields(14)), "###,###,###,###,##0.00") & ") 點"
        Printer.CurrentX = 0
        Printer.CurrentY = iPrint
        'edit by nickc 2005/05/04
        'Printer.Print "超過承辦期限之件數：" & Format(IIf(CheckStr(.Fields(7)) = "", "0", CheckStr(.Fields(7))), "###,###,###,###,##0") & " 件"
        Printer.Print "超過承辦期限之件數：" & Format("0" & CheckStr(.Fields(7)), "###,###,###,###,##0") & "(" & Format("0" & CheckStr(.Fields(15)), "###,###,###,###,##0.00") & ") 件"
        iPrint = iPrint + 300
        If iPrint >= 9000 Then
            Page = Page + 1
            Printer.NewPage
            PrintTitle_2
        End If
        ShowLine
    End If
End With
CheckOC
End Sub

Sub PrintTitle_2() '列印抬頭
frm090614.Hide
iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "工作進度資料表(專利)"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
'If Len(frm090614.txt1(3)) <> 0 And Len(frm090614.txt1(4)) <> 0 Then
'   Printer.Print "發文年月：" & frm090614.txt1(3) & "/" & frm090614.txt1(4)
'Else
'   Printer.Print "發文年月：" & str(Val(Mid(GetTodayDate, 1, 4)) - 1911) & "/" & Mid(GetTodayDate, 5, 2)
'End If
Printer.Print "發文年月：" & Me.m_strYear & "/" & Me.m_strMonth
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "承辦人：" & GetPrjSalesNM(strTemp3)
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
GetPleft_2
Printer.Font.Size = 9
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "目次"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "收文"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "收文日"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "Y/N"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "本所"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "法定"
Printer.CurrentX = PLeft(12)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "會稿日"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "核稿人"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "會稿"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "承辦"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "備註"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "類別"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "期限"
Printer.CurrentX = PLeft(16)
Printer.CurrentY = iPrint
Printer.Print "完成日"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "天數"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle_2
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090201_3_2 = Nothing
End Sub
