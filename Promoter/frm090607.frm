VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090607 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員收文高低標查詢"
   ClientHeight    =   3390
   ClientLeft      =   435
   ClientTop       =   405
   ClientWidth     =   4080
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   4080
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   2820
      TabIndex        =   11
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2040
      TabIndex        =   10
      Top             =   20
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   10
      Left            =   1164
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2985
      Width           =   300
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   9
      Left            =   1164
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2622
      Width           =   330
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   8
      Left            =   1164
      MaxLength       =   1
      TabIndex        =   7
      Top             =   2263
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   7
      Left            =   1164
      MaxLength       =   6
      TabIndex        =   6
      Top             =   1904
      Width           =   885
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   6
      Left            =   2184
      MaxLength       =   3
      TabIndex        =   5
      Top             =   1545
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   5
      Left            =   1164
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1545
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   4
      Left            =   2190
      MaxLength       =   5
      TabIndex        =   3
      Top             =   1186
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   3
      Left            =   1164
      MaxLength       =   5
      TabIndex        =   2
      Top             =   1186
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   1
      Left            =   1176
      MaxLength       =   4
      TabIndex        =   1
      Top             =   827
      Width           =   900
   End
   Begin VB.TextBox txt1 
      Height          =   300
      Index           =   0
      Left            =   1164
      MaxLength       =   3
      TabIndex        =   0
      Top             =   468
      Width           =   900
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   1
      Left            =   2130
      TabIndex        =   26
      Top             =   850
      Width           =   1890
      VariousPropertyBits=   27
      Size            =   "3334;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   255
      Index           =   0
      Left            =   2130
      TabIndex        =   25
      Top             =   1927
      Width           =   1500
      VariousPropertyBits=   27
      Size            =   "2646;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "請自行輸入系統別"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   2136
      TabIndex        =   23
      Top             =   504
      Width           =   1512
   End
   Begin VB.Line Line3 
      X1              =   1668
      X2              =   2823
      Y1              =   1688
      Y2              =   1688
   End
   Begin VB.Line Line2 
      X1              =   1716
      X2              =   2871
      Y1              =   1329
      Y2              =   1329
   End
   Begin VB.Label Label1 
      Caption         =   "(1.各區  2.合併)"
      Height          =   180
      Index           =   10
      Left            =   1515
      TabIndex        =   22
      Top             =   3045
      Width           =   1515
   End
   Begin VB.Label Label1 
      Caption         =   "(1.螢幕  2.報表)"
      Height          =   180
      Index           =   9
      Left            =   1524
      TabIndex        =   21
      Top             =   2682
      Width           =   1632
   End
   Begin VB.Label Label1 
      Caption         =   "(1.新申請案  2.爭議/救濟案)"
      Height          =   180
      Index           =   8
      Left            =   1536
      TabIndex        =   20
      Top             =   2323
      Width           =   2472
   End
   Begin VB.Label Label1 
      Caption         =   "顯示內容："
      Height          =   180
      Index           =   7
      Left            =   180
      TabIndex        =   19
      Top             =   3045
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "顯示方式："
      Height          =   180
      Index           =   6
      Left            =   180
      TabIndex        =   18
      Top             =   2682
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "查詢性質："
      Height          =   180
      Index           =   5
      Left            =   180
      TabIndex        =   17
      Top             =   2323
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "智權人員："
      Height          =   180
      Index           =   4
      Left            =   180
      TabIndex        =   16
      Top             =   1964
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "業務區："
      Height          =   180
      Index           =   3
      Left            =   180
      TabIndex        =   15
      Top             =   1605
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "收文年月："
      Height          =   180
      Index           =   2
      Left            =   180
      TabIndex        =   14
      Top             =   1246
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      Height          =   180
      Index           =   1
      Left            =   180
      TabIndex        =   13
      Top             =   887
      Width           =   1008
   End
   Begin VB.Label Label1 
      Caption         =   "系統類別："
      Height          =   180
      Index           =   0
      Left            =   180
      TabIndex        =   12
      Top             =   504
      Width           =   1008
   End
End
Attribute VB_Name = "frm090607"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; lbl1(index) ; Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/16 日期欄已修改
Option Explicit
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, L(1 To 5, 1 To 12) As Single
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 17) As String, strTemp3 As String, IngL(1 To 3) As Single, tmpnickG As Integer, k As Integer
Dim PLeft(0 To 13) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, StrTemp99(0 To 7) As String, StrTemp7(0 To 9) As String
Dim allG As Integer, Hightmp As Single, Middletmp As Single, Lowtmp As Single, Avgtmp As Single, Sumtmp As Single, SumCount As Single
Dim RsTmpNick As New ADODB.Recordset

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     If Len(txt1(0)) = 0 Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     End If
      If Len(txt1(1)) = 0 Then
         s = MsgBox("申請國家不可空白!!", , "USER 輸入錯誤")
         txt1(1).SetFocus
         Exit Sub
      End If
      If Len(txt1(3)) = 0 Or Len(txt1(4)) = 0 Then
          s = MsgBox("收文年月區間不可空白!!", , "USER 輸入錯誤")
          If Len(txt1(4)) = 0 Then txt1(4).SetFocus
          If Len(txt1(3)) = 0 Then txt1(3).SetFocus
          Exit Sub
      End If
      If Len(txt1(8)) = 0 Then
          s = MsgBox("查詢性質不可空白!!", , "USER 輸入錯誤")
          txt1(8).SetFocus
          Exit Sub
      End If
      If Len(txt1(9)) = 0 Then
          s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
          txt1(9).SetFocus
          Exit Sub
      End If
      If Len(txt1(10)) = 0 Then
          s = MsgBox("顯示內容不可空白!!", , "USER 輸入錯誤")
          txt1(10).SetFocus
          Exit Sub
      End If
         'Add By Cheng 2002/03/21
      If PUB_CheckKeyInYYMM(Me.txt1(3)) = -1 Then
         Me.txt1(3).SetFocus
         txt1_GotFocus 3
         Exit Sub
      End If
      If PUB_CheckKeyInYYMM(Me.txt1(4)) = -1 Then
         Me.txt1(4).SetFocus
         txt1_GotFocus 4
         Exit Sub
      End If
      Screen.MousePointer = vbHourglass
      Me.Enabled = False
      ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/20 清除查詢印表記錄檔欄位
      'If StrTemp99(0) <> txt1(0) Or StrTemp99(1) <> txt1(1) Or StrTemp99(3) <> txt1(3) Or StrTemp99(4) <> txt1(4) Or StrTemp99(5) <> txt1(5) Or StrTemp99(6) <> txt1(6) Or StrTemp99(7) <> txt1(7) Then
          Process2
      '    StrTemp99(0) = txt1(0)
      '    StrTemp99(1) = txt1(1)
      '    StrTemp99(3) = txt1(3)
      '    StrTemp99(4) = txt1(4)
      '    StrTemp99(5) = txt1(5)
      '    StrTemp99(6) = txt1(6)
      '    StrTemp99(7) = txt1(7)
      'End If
      Process
      Me.Enabled = True
      Screen.MousePointer = vbDefault
Case 1
     Unload Me
Case Else
End Select
End Sub

'91.08.07 nick 大幅修改
Sub Process1()
cnnConnection.Execute "DELETE FROM R090607_3 WHERE ID='" & strUserNum & "' "
If Val(txt1(8)) = 1 Then
    pub_QL05 = pub_QL05 & ";" & Label1(5) & "1.新申請案" 'Add By Sindy 2010/12/20
'新申請案
    '91.08.07   nick  新申請案
    '計算差
    '合計件
    CheckOC
    strSql = "select aaa.r099001 as r099001,aaa.r099018 as r099018," & _
             "sum(nvl(aa1.r099003,0)) as r099003,sum(nvl(aa2.r099004,0)) as r099004,sum(nvl(aa3.r099005,0)) as r099005,sum(nvl(aa4.r099006,0)) as r099006,sum(nvl(aa5.r099007,0)) as r099007," & _
             "sum(nvl(aa6.r099008,0)) as r099008,sum(nvl(aa7.r099009,0)) as r099009,sum(nvl(aa8.r099010,0)) as r099010,sum(nvl(aa9.r099011,0)) as r099011,sum(nvl(aa10.r099012,0)) as r099012," & _
             "Sum(nvl(aa11.r099013,0)) as r099013, Sum(nvl(aa12.r099014, 0)) as r099014, Sum(nvl(aa13.r099015, 0)) as r099015, Sum(nvl(aa14.r099016, 0)) as r099016, Sum(nvl(aa15.r099017, 0)) as r099017," & _
             "sum(nvl(aa1.S099003,0)) as S099003,sum(nvl(aa2.S099004,0)) as S099004,sum(nvl(aa3.S099005,0)) as S099005,sum(nvl(aa4.S099006,0)) as S099006,sum(nvl(aa5.S099007,0)) as S099007,"
    strSql = strSql & "sum(nvl(aa6.S099008,0)) as S099008,sum(nvl(aa7.S099009,0)) as S099009,sum(nvl(aa8.S099010,0)) as S099010,sum(nvl(aa9.S099011,0)) as S099011,sum(nvl(aa10.S099012,0)) as S099012,"
    strSql = strSql & "sum(nvl(aa11.S099013,0)) as S099013,sum(nvl(aa12.S099014,0)) as S099014,sum(nvl(aa13.S099015,0)) as S099015,sum(nvl(aa14.S099016,0)) as S099016,sum(nvl(aa15.S099017,0)) as S099017 "
    strSql = strSql & " From " & _
             " (select distinct r099001 as r099001,substr(R099018,1,6)||'01' as r099018 from r090607_1 where id='" & strUserNum & "') aaa," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099003) as r099003,sum(r099003) as S099003 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099003 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa1," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099004) as r099004,sum(r099004) as S099004 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099004 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa2," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099005) as r099005,sum(r099005) as S099005 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099005 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa3," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099006) as r099006,sum(r099006) as S099006 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099006 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa4," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099007) as r099007,sum(r099007) as S099007 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099007 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa5," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099008) as r099008,sum(r099008) as S099008 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099008 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa6," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099009) as r099009,sum(r099009) as S099009 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099009 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa7," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099010) as r099010,sum(r099010) as S099010 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099010 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa8," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099011) as r099011,sum(r099011) as S099011 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099011 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa9," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099012) as r099012,sum(r099012) as S099012 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099012 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa10," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099013) as r099013,sum(r099013) as S099013 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099013 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa11," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099014) as r099014,sum(r099014) as S099014 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099014 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa12," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099015) as r099015,sum(r099015) as S099015 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099015 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa13," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099016) as r099016,sum(r099016) as S099016 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099016 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa14," & _
             " (SELECT R099001 as r099001,substr(R099018,1,6)||'01' as r099018,COUNT(R099017) as r099017,sum(r099017) as S099017 FROM R090607_1 WHERE ID='" & strUserNum & "' and r099017 <> 0 GROUP BY R099001,substr(R099018,1,6)||'01') aa15 " & _
             " where aaa.r099001=aa1.r099001(+) and aaa.r099018=aa1.r099018(+) and aaa.r099001=aa2.r099001(+) and aaa.r099018=aa2.r099018(+)" & _
             " and aaa.r099001=aa3.r099001(+) and aaa.r099018=aa3.r099018(+) and aaa.r099001=aa4.r099001(+) and aaa.r099018=aa4.r099018(+)" & _
             " and aaa.r099001=aa5.r099001(+) and aaa.r099018=aa5.r099018(+) and aaa.r099001=aa6.r099001(+) and aaa.r099018=aa6.r099018(+)" & _
             " and aaa.r099001=aa7.r099001(+) and aaa.r099018=aa7.r099018(+) and aaa.r099001=aa8.r099001(+) and aaa.r099018=aa8.r099018(+)"
    strSql = strSql & " and aaa.r099001=aa9.r099001(+) and aaa.r099018=aa9.r099018(+) and aaa.r099001=aa10.r099001(+) and aaa.r099018=aa10.r099018(+)" & _
             " and aaa.r099001=aa11.r099001(+) and aaa.r099018=aa11.r099018(+) and aaa.r099001=aa12.r099001(+) and aaa.r099018=aa12.r099018(+)" & _
             " and aaa.r099001=aa13.r099001(+) and aaa.r099018=aa13.r099018(+) and aaa.r099001=aa14.r099001(+) and aaa.r099018=aa14.r099018(+)" & _
             " and aaa.r099001=aa15.r099001(+) and aaa.r099018=aa15.r099018(+) group by aaa.r099001,aaa.r099018 "
    With adoRecordset
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While .EOF = False
                Erase L
                '合計件
                '高標
                    L(1, 1) = Val(CheckStr(.Fields("r099003").Value))   '發明
                    L(1, 2) = Val(CheckStr(.Fields("r099004").Value))   '新型
                    L(1, 3) = Val(CheckStr(.Fields("r099005").Value))   '設計
                    L(1, 4) = Val(CheckStr(.Fields("r099006").Value))   '再審
                '介於
                    L(1, 5) = Val(CheckStr(.Fields("r099008").Value))   '發明
                    L(1, 6) = Val(CheckStr(.Fields("r099009").Value))   '新型
                    L(1, 7) = Val(CheckStr(.Fields("r099010").Value))   '設計
                    L(1, 8) = Val(CheckStr(.Fields("r099011").Value))   '再審
                '低標
                    L(1, 9) = Val(CheckStr(.Fields("r099013").Value))   '發明
                    L(1, 10) = Val(CheckStr(.Fields("r099014").Value))  '新型
                    L(1, 11) = Val(CheckStr(.Fields("r099015").Value))  '設計
                    L(1, 12) = Val(CheckStr(.Fields("r099016").Value))  '再審
                '合計點
                '高標
                    L(2, 1) = Val(Format(CheckStr(.Fields("S099003").Value), "####.00"))   '發明
                    L(2, 2) = Val(Format(CheckStr(.Fields("S099004").Value), "####.00"))   '新型
                    L(2, 3) = Val(Format(CheckStr(.Fields("S099005").Value), "####.00"))   '設計
                    L(2, 4) = Val(Format(CheckStr(.Fields("S099006").Value), "####.00"))   '再審
                '介於
                    L(2, 5) = Val(Format(CheckStr(.Fields("S099008").Value), "####.00"))   '發明
                    L(2, 6) = Val(Format(CheckStr(.Fields("S099009").Value), "####.00"))   '新型
                    L(2, 7) = Val(Format(CheckStr(.Fields("S099010").Value), "####.00"))   '設計
                    L(2, 8) = Val(Format(CheckStr(.Fields("S099011").Value), "####.00"))   '再審
                '低標
                    L(2, 9) = Val(Format(CheckStr(.Fields("S099013").Value), "####.00"))   '發明
                    L(2, 10) = Val(Format(CheckStr(.Fields("S099014").Value), "####.00"))  '新型
                    L(2, 11) = Val(Format(CheckStr(.Fields("S099015").Value), "####.00"))  '設計
                    L(2, 12) = Val(Format(CheckStr(.Fields("S099016").Value), "####.00"))  '再審
                '平均點
                '高標
                    If L(2, 1) <> 0 Then    '發明
                        L(3, 1) = Val(Format(Trim(L(2, 1) / L(1, 1)), "####.00"))
                    Else
                        L(3, 1) = 0
                    End If
                    If L(2, 2) <> 0 Then    '新型
                        L(3, 2) = Format(Trim(L(2, 2) / L(1, 2)), "####.00")
                    Else
                        L(3, 2) = 0
                    End If
                    If L(2, 3) <> 0 Then    '設計
                        L(3, 3) = Format(Trim(L(2, 3) / L(1, 3)), "####.00")
                    Else
                        L(3, 3) = 0
                    End If
                    If L(2, 4) <> 0 Then    '再審
                        L(3, 4) = Format(Trim(L(2, 4) / L(1, 4)), "####.00")
                    Else
                        L(3, 4) = 0
                    End If
                '介於
                    If L(2, 5) <> 0 Then
                        L(3, 5) = Format(Trim(L(2, 5) / L(1, 5)), "####.00")  '發明
                    Else
                        L(3, 5) = 0
                    End If
                    If L(2, 6) <> 0 Then
                        L(3, 6) = Format(Trim(L(2, 6) / L(1, 6)), "####.00")  '新型
                    Else
                        L(3, 6) = 0
                    End If
                    If L(2, 7) <> 0 Then
                        L(3, 7) = Format(Trim(L(2, 7) / L(1, 7)), "####.00") '設計
                    Else
                        L(3, 7) = 0
                    End If
                    If L(2, 8) <> 0 Then
                        L(3, 8) = Format(Trim(L(2, 8) / L(1, 8)), "####.00")  '再審
                    Else
                        L(3, 8) = 0
                    End If
                '低標
                    If L(2, 9) <> 0 Then
                        L(3, 9) = Format(Trim(L(2, 9) / L(1, 9)), "####.00")  '發明
                    Else
                        L(3, 9) = 0
                    End If
                    If L(2, 10) <> 0 Then
                        L(3, 10) = Format(Trim(L(2, 10) / L(1, 10)), "####.00") '新型
                    Else
                        L(3, 10) = 0
                    End If
                    If L(2, 11) <> 0 Then
                        L(3, 11) = Format(Trim(L(2, 11) / L(1, 11)), "####.00") '設計
                    Else
                        L(3, 11) = 0
                    End If
                    If L(2, 12) <> 0 Then
                        L(3, 12) = Format(Trim(L(2, 12) / L(1, 12)), "####.00") '再審
                    Else
                        L(3, 12) = 0
                    End If
                '標準價
                L(4, 1) = 0
                L(4, 2) = 0
                L(4, 3) = 0
                L(4, 4) = 0
                L(4, 5) = 0
                L(4, 6) = 0
                L(4, 7) = 0
                L(4, 8) = 0
                L(4, 9) = 0
                L(4, 10) = 0
                L(4, 11) = 0
                L(4, 12) = 0
                'Modify By Cheng 2003/07/15
'                strSQL = "select cf03,cf13 from casefee where cf01='" & Txt1(0) & "' and cf02='" & Txt1(1) & "' and cf03 in ('101','102','103','107') "
'                Set RsTmpNick = New ADODB.Recordset
'                RsTmpNick.CursorLocation = adUseClient
'                RsTmpNick.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'                If RsTmpNick.RecordCount <> 0 Then
'                    RsTmpNick.MoveFirst
'                    Do While RsTmpNick.EOF = False
'                        If CheckStr(RsTmpNick.Fields("cf03").Value) = "101" Then  '發明
'                            L(4, 1) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 5) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 9) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                        End If
'                        If CheckStr(RsTmpNick.Fields("cf03").Value) = "102" Then  '新型
'                            L(4, 2) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 6) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 10) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                        End If
'                        If CheckStr(RsTmpNick.Fields("cf03").Value) = "103" Then  '設計
'                            L(4, 3) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 7) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 11) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                        End If
'                        If CheckStr(RsTmpNick.Fields("cf03").Value) = "107" Then  '再審
'                            L(4, 4) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 8) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                            L(4, 12) = Val(CheckStr(RsTmpNick.Fields("cf13").Value))
'                        End If
'                        RsTmpNick.MoveNext
'                    Loop
'                End If
                strSql = "Select Sum(Decode(R099003,0,0,Nvl(CP33,0))), Sum(Decode(R099003,0,0,Nvl(1,0))), Sum(Decode(R099008,0,0,Nvl(CP33,0))), Sum(Decode(R099008,0,0,Nvl(1,0))), Sum(Decode(R099013,0,0,Nvl(CP33,0))), Sum(Decode(R099013,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='101' Or (CP10='104' And PA08='1' )) " & IIf(IsNull(.Fields("R099001").Value) = False, " And R099001='" & .Fields("R099001").Value & "'", " And R099001 Is Null") & " And substr(R099018,1,6)=" & Mid("" & .Fields("R099018").Value, 1, 6)
                Set RsTmpNick = New ADODB.Recordset
                RsTmpNick.CursorLocation = adUseClient
                RsTmpNick.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If RsTmpNick.RecordCount <> 0 Then
                    If Val("" & RsTmpNick.Fields(1).Value) <> 0 Then
                        L(4, 1) = Format(Val(CheckStr(RsTmpNick.Fields(0))) / Val(RsTmpNick.Fields(1).Value), "#########0.00")
                    Else
                        L(4, 1) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(3).Value) <> 0 Then
                        L(4, 5) = Format(Val(CheckStr(RsTmpNick.Fields(2))) / Val(RsTmpNick.Fields(3).Value), "#########0.00")
                    Else
                        L(4, 5) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(5).Value) <> 0 Then
                        L(4, 9) = Format(Val(CheckStr(RsTmpNick.Fields(4))) / Val(RsTmpNick.Fields(5).Value), "#########0.00")
                    Else
                        L(4, 9) = Format(0, "#########0.00")
                    End If
                End If
                If RsTmpNick.State <> adStateClosed Then RsTmpNick.Close
                Set RsTmpNick = Nothing
                strSql = "Select Sum(Decode(R099004,0,0,Nvl(CP33,0))), Sum(Decode(R099004,0,0,Nvl(1,0))), Sum(Decode(R099009,0,0,Nvl(CP33,0))), Sum(Decode(R099009,0,0,Nvl(1,0))), Sum(Decode(R099014,0,0,Nvl(CP33,0))), Sum(Decode(R099014,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='102' Or (CP10='104' And PA08='2' )) " & IIf(IsNull(.Fields("R099001").Value) = False, " And R099001='" & .Fields("R099001").Value & "'", " And R099001 Is Null") & " And substr(R099018,1,6)=" & Mid("" & .Fields("R099018").Value, 1, 6)
                Set RsTmpNick = New ADODB.Recordset
                RsTmpNick.CursorLocation = adUseClient
                RsTmpNick.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If RsTmpNick.RecordCount <> 0 Then
                    If Val("" & RsTmpNick.Fields(1).Value) <> 0 Then
                        L(4, 2) = Format(Val(CheckStr(RsTmpNick.Fields(0))) / Val(RsTmpNick.Fields(1).Value), "#########0.00")
                    Else
                        L(4, 2) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(3).Value) <> 0 Then
                        L(4, 6) = Format(Val(CheckStr(RsTmpNick.Fields(2))) / Val(RsTmpNick.Fields(3).Value), "#########0.00")
                    Else
                        L(4, 6) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(5).Value) <> 0 Then
                        L(4, 10) = Format(Val(CheckStr(RsTmpNick.Fields(4))) / Val(RsTmpNick.Fields(5).Value), "#########0.00")
                    Else
                        L(4, 10) = Format(0, "#########0.00")
                    End If
                End If
                If RsTmpNick.State <> adStateClosed Then RsTmpNick.Close
                Set RsTmpNick = Nothing
                strSql = "Select Sum(Decode(R099005,0,0,Nvl(CP33,0))), Sum(Decode(R099005,0,0,Nvl(1,0))), Sum(Decode(R099010,0,0,Nvl(CP33,0))), Sum(Decode(R099010,0,0,Nvl(1,0))), Sum(Decode(R099015,0,0,Nvl(CP33,0))), Sum(Decode(R099015,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='103' Or CP10='105' Or CP10='125') " & IIf(IsNull(.Fields("R099001").Value) = False, " And R099001='" & .Fields("R099001").Value & "'", " And R099001 Is Null") & " And substr(R099018,1,6)=" & Mid("" & .Fields("R099018").Value, 1, 6)
                Set RsTmpNick = New ADODB.Recordset
                RsTmpNick.CursorLocation = adUseClient
                RsTmpNick.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If RsTmpNick.RecordCount <> 0 Then
                    If Val("" & RsTmpNick.Fields(1).Value) <> 0 Then
                        L(4, 3) = Format(Val(CheckStr(RsTmpNick.Fields(0))) / Val(RsTmpNick.Fields(1).Value), "#########0.00")
                    Else
                        L(4, 3) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(3).Value) <> 0 Then
                        L(4, 7) = Format(Val(CheckStr(RsTmpNick.Fields(2))) / Val(RsTmpNick.Fields(3).Value), "#########0.00")
                    Else
                        L(4, 7) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(5).Value) <> 0 Then
                        L(4, 11) = Format(Val(CheckStr(RsTmpNick.Fields(4))) / Val(RsTmpNick.Fields(5).Value), "#########0.00")
                    Else
                        L(4, 11) = Format(0, "#########0.00")
                    End If
                End If
                If RsTmpNick.State <> adStateClosed Then RsTmpNick.Close
                Set RsTmpNick = Nothing
                strSql = "Select Sum(Decode(R099006,0,0,Nvl(CP33,0))), Sum(Decode(R099006,0,0,Nvl(1,0))), Sum(Decode(R099011,0,0,Nvl(CP33,0))), Sum(Decode(R099011,0,0,Nvl(1,0))), Sum(Decode(R099016,0,0,Nvl(CP33,0))), Sum(Decode(R099016,0,0,Nvl(1,0))) FROM R090607_1, Caseprogress, Patent WHERE R099020=CP09 And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04 AND (CP10='107' ) " & IIf(IsNull(.Fields("R099001").Value) = False, " And R099001='" & .Fields("R099001").Value & "'", " And R099001 Is Null") & " And substr(R099018,1,6)=" & Mid("" & .Fields("R099018").Value, 1, 6)
                Set RsTmpNick = New ADODB.Recordset
                RsTmpNick.CursorLocation = adUseClient
                RsTmpNick.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                If RsTmpNick.RecordCount <> 0 Then
                    If Val("" & RsTmpNick.Fields(1).Value) <> 0 Then
                        L(4, 4) = Format(Val(CheckStr(RsTmpNick.Fields(0))) / Val(RsTmpNick.Fields(1).Value), "#########0.00")
                    Else
                        L(4, 4) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(3).Value) <> 0 Then
                        L(4, 8) = Format(Val(CheckStr(RsTmpNick.Fields(2))) / Val(RsTmpNick.Fields(3).Value), "#########0.00")
                    Else
                        L(4, 8) = Format(0, "#########0.00")
                    End If
                    If Val("" & RsTmpNick.Fields(5).Value) <> 0 Then
                        L(4, 12) = Format(Val(CheckStr(RsTmpNick.Fields(4))) / Val(RsTmpNick.Fields(5).Value), "#########0.00")
                    Else
                        L(4, 12) = Format(0, "#########0.00")
                    End If
                End If
                If RsTmpNick.State <> adStateClosed Then RsTmpNick.Close
                Set RsTmpNick = Nothing
                '差距
                '高標
                     L(5, 1) = Format(IIf(L(3, 1) = 0, 0, L(3, 1) - L(4, 1)), "####.00") '發明
                     L(5, 2) = Format(IIf(L(3, 2) = 0, 0, L(3, 2) - L(4, 2)), "####.00") '新型
                     L(5, 3) = Format(IIf(L(3, 3) = 0, 0, L(3, 3) - L(4, 3)), "####.00") '設計
                     L(5, 4) = Format(IIf(L(3, 4) = 0, 0, L(3, 4) - L(4, 4)), "####.00") '再審
                '介於
                     L(5, 5) = Format(IIf(L(3, 5) = 0, 0, L(3, 5) - L(4, 5)), "####.00") '發明
                     L(5, 6) = Format(IIf(L(3, 6) = 0, 0, L(3, 6) - L(4, 6)), "####.00") '新型
                     L(5, 7) = Format(IIf(L(3, 7) = 0, 0, L(3, 7) - L(4, 7)), "####.00") '設計
                     L(5, 8) = Format(IIf(L(3, 8) = 0, 0, L(3, 8) - L(4, 8)), "####.00") '再審
                '低標
                     L(5, 9) = Format(IIf(L(3, 9) = 0, 0, L(3, 9) - L(4, 9)), "####.00") '發明
                     L(5, 10) = Format(IIf(L(3, 10) = 0, 0, L(3, 10) - L(4, 10)), "####.00") '新型
                     L(5, 11) = Format(IIf(L(3, 11) = 0, 0, L(3, 11) - L(4, 11)), "####.00") '設計
                     L(5, 12) = Format(IIf(L(3, 12) = 0, 0, L(3, 12) - L(4, 12)), "####.00") '再審
                '高標差
                Hightmp = Format(Trim(str((L(1, 1) * L(5, 1)) + (L(1, 2) * L(5, 2)) + (L(1, 3) * L(5, 3)) + (L(1, 4) * L(5, 4)))), "####.00")
                '介於差
                Middletmp = Format(Trim(str((L(1, 5) * L(5, 5)) + (L(1, 6) * L(5, 6)) + (L(1, 7) * L(5, 7)) + (L(1, 8) * L(5, 8)))), "####.00")
                '低標差
                Lowtmp = Format(Trim(str((L(1, 9) * L(5, 9)) + (L(1, 10) * L(5, 10)) + (L(1, 11) * L(5, 11)) + (L(1, 12) * L(5, 12)))), "####.00")
                '總和
                Sumtmp = Hightmp + Middletmp + Lowtmp
                '高標，介於，低標合計件
                SumCount = L(1, 1) + L(1, 2) + L(1, 3) + L(1, 4) + L(1, 5) + L(1, 6) + L(1, 7) + L(1, 8) + L(1, 9) + L(1, 10) + L(1, 11) + L(1, 12)
                '平均差
                If SumCount <> 0 Then
                    Avgtmp = Format(IIf(SumCount <> 0, Sumtmp / SumCount, 0), "###.00")
                Else
                    Avgtmp = 0
                End If
                '存入暫存檔
                strSql = "insert into r090607_3 (r101001,r101002,r101003,r101004,r101005,r101006,r101007,r101008,id) values ('" & CheckStr(.Fields("r099001").Value) & "','" & CheckStr(.Fields("r099018").Value) & "'," & _
                         L(1, 1) + L(1, 2) + L(1, 3) + L(1, 4) & "," & L(1, 5) + L(1, 6) + L(1, 7) + L(1, 8) & "," & L(1, 9) + L(1, 10) + L(1, 11) + L(1, 12) & "," & Middletmp & "," & Sumtmp & "," & Avgtmp & ",'" & strUserNum & "')"
                cnnConnection.Execute strSql
                .MoveNext
            Loop
        End If
    End With
    CheckOC
Else
    pub_QL05 = pub_QL05 & ";" & Label1(5) & "2.爭議/救濟案" 'Add By Sindy 2010/12/20
'爭議救濟案
    '91.08.07   nick  爭議救濟案
    '計算差
    '合計件
    CheckOC
    strSql = "select aaa.r100001 as r100001,aaa.r100015 as r100015," & _
             "sum(nvl(aa1.r100003,0)) as r100003,sum(nvl(aa2.r100004,0)) as r100004,sum(nvl(aa3.r100005,0)) as r100005,sum(nvl(aa4.r100006,0)) as r100006," & _
             "sum(nvl(aa5.r100007,0)) as r100007,sum(nvl(aa6.r100008,0)) as r100008,sum(nvl(aa7.r100009,0)) as r100009,sum(nvl(aa8.r100010,0)) as r100010," & _
             "sum(nvl(aa9.r100011,0)) as r100011,sum(nvl(aa10.r100012,0)) as r100012,sum(nvl(aa11.r100013,0)) as r100013,sum(nvl(aa12.r100014,0)) as r100014," & _
             "sum(nvl(aa1.s100003,0)) as s100003,sum(nvl(aa2.s100004,0)) as s100004,sum(nvl(aa3.s100005,0)) as s100005,sum(nvl(aa4.s100006,0)) as s100006," & _
             "sum(nvl(aa5.s100007,0)) as s100007,sum(nvl(aa6.s100008,0)) as s100008,sum(nvl(aa7.s100009,0)) as s100009,sum(nvl(aa8.s100010,0)) as s100010," & _
             "sum(nvl(aa9.s100011,0)) as s100011,sum(nvl(aa10.s100012,0)) as s100012,sum(nvl(aa11.s100013,0)) as s100013,sum(nvl(aa12.s100014,0)) as s100014" & _
             " From " & _
             " (select distinct r100001 as r100001,substr(r100015,1,6)||'01' as r100015 from r090607_2 where id='" & strUserNum & "') aaa, " & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100003) as r100003,sum(r100003) as s100003 from r090607_2 where id='" & strUserNum & "' and r100003<>0 group by r100001,substr(r100015,1,6)||'01') aa1," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100004) as r100004,sum(r100004) as s100004 from r090607_2 where id='" & strUserNum & "' and r100004<>0 group by r100001,substr(r100015,1,6)||'01') aa2," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100005) as r100005,sum(r100005) as s100005 from r090607_2 where id='" & strUserNum & "' and r100005<>0 group by r100001,substr(r100015,1,6)||'01') aa3," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100006) as r100006,sum(r100006) as s100006 from r090607_2 where id='" & strUserNum & "' and r100006<>0 group by r100001,substr(r100015,1,6)||'01') aa4," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100007) as r100007,sum(r100007) as s100007 from r090607_2 where id='" & strUserNum & "' and r100007<>0 group by r100001,substr(r100015,1,6)||'01') aa5," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100008) as r100008,sum(r100008) as s100008 from r090607_2 where id='" & strUserNum & "' and r100008<>0 group by r100001,substr(r100015,1,6)||'01') aa6," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100009) as r100009,sum(r100009) as s100009 from r090607_2 where id='" & strUserNum & "' and r100009<>0 group by r100001,substr(r100015,1,6)||'01') aa7," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100010) as r100010,sum(r100010) as s100010 from r090607_2 where id='" & strUserNum & "' and r100010<>0 group by r100001,substr(r100015,1,6)||'01') aa8," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100011) as r100011,sum(r100011) as s100011 from r090607_2 where id='" & strUserNum & "' and r100011<>0 group by r100001,substr(r100015,1,6)||'01') aa9," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100012) as r100012,sum(r100012) as s100012 from r090607_2 where id='" & strUserNum & "' and r100012<>0 group by r100001,substr(r100015,1,6)||'01') aa10," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100013) as r100013,sum(r100013) as s100013 from r090607_2 where id='" & strUserNum & "' and r100013<>0 group by r100001,substr(r100015,1,6)||'01') aa11," & _
             " (select r100001 as r100001,substr(r100015,1,6)||'01' as r100015,count(r100014) as r100014,sum(r100014) as s100014 from r090607_2 where id='" & strUserNum & "' and r100014<>0 group by r100001,substr(r100015,1,6)||'01') aa12 " & _
             " where aaa.r100001=aa1.r100001(+) and aaa.r100015=aa1.r100015(+) and aaa.r100001=aa2.r100001(+) and aaa.r100015=aa2.r100015(+) " & _
             " and aaa.r100001=aa3.r100001(+) and aaa.r100015=aa3.r100015(+) and aaa.r100001=aa4.r100001(+) and aaa.r100015=aa4.r100015(+) " & _
             " and aaa.r100001=aa5.r100001(+) and aaa.r100015=aa5.r100015(+) and aaa.r100001=aa6.r100001(+) and aaa.r100015=aa6.r100015(+) " & _
             " and aaa.r100001=aa7.r100001(+) and aaa.r100015=aa7.r100015(+) and aaa.r100001=aa8.r100001(+) and aaa.r100015=aa8.r100015(+) "
    strSql = strSql & " and aaa.r100001=aa9.r100001(+) and aaa.r100015=aa9.r100015(+) and aaa.r100001=aa10.r100001(+) and aaa.r100015=aa10.r100015(+) " & _
             " and aaa.r100001=aa11.r100001(+) and aaa.r100015=aa11.r100015(+) and aaa.r100001=aa12.r100001(+) and aaa.r100015=aa12.r100015(+) " & _
             " group by aaa.r100001,aaa.r100015"
    With adoRecordset
        .CursorLocation = adUseClient
        .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If .RecordCount <> 0 Then
            .MoveFirst
            Do While .EOF = False
                Erase L
                '合計件
                '高標
                    L(1, 1) = Val(CheckStr(.Fields("r100003").Value))   '爭議
                    L(1, 2) = Val(CheckStr(.Fields("r100004").Value))   '救濟
                    L(1, 3) = Val(CheckStr(.Fields("r100005").Value))   '其他
                    L(1, 4) = 0   '空白  暫時無用
                '介於
                    L(1, 5) = Val(CheckStr(.Fields("r100007").Value))   '爭議
                    L(1, 6) = Val(CheckStr(.Fields("r100008").Value))   '救濟
                    L(1, 7) = Val(CheckStr(.Fields("r100009").Value))   '其他
                    L(1, 8) = 0   '空白  暫時無用
                '低標
                    L(1, 9) = Val(CheckStr(.Fields("r100011").Value))   '爭議
                    L(1, 10) = Val(CheckStr(.Fields("r100012").Value))  '救濟
                    L(1, 11) = Val(CheckStr(.Fields("r100013").Value))  '其他
                    L(1, 12) = 0  '空白  暫時無用
                '合計點
                '高標
                    L(2, 1) = Val(Format(CheckStr(.Fields("S100003").Value), "####.00"))   '爭議
                    L(2, 2) = Val(Format(CheckStr(.Fields("S100004").Value), "####.00"))   '救濟
                    L(2, 3) = Val(Format(CheckStr(.Fields("S100005").Value), "####.00"))   '其他
                    L(2, 4) = 0   '空白  暫時無用
                '介於
                    L(2, 5) = Val(Format(CheckStr(.Fields("S100007").Value), "####.00"))   '爭議
                    L(2, 6) = Val(Format(CheckStr(.Fields("S100008").Value), "####.00"))   '救濟
                    L(2, 7) = Val(Format(CheckStr(.Fields("S100009").Value), "####.00"))   '其他
                    L(2, 8) = 0   '空白  暫時無用
                '低標
                    L(2, 9) = Val(Format(CheckStr(.Fields("S100011").Value), "####.00"))   '爭議
                    L(2, 10) = Val(Format(CheckStr(.Fields("S100012").Value), "####.00"))  '救濟
                    L(2, 11) = Val(Format(CheckStr(.Fields("S100013").Value), "####.00"))  '其他
                    L(2, 12) = 0  '空白  暫時無用
                '平均點
                '高標
                    If L(2, 1) <> 0 Then    '爭議
                        L(3, 1) = Val(Format(Trim(L(2, 1) / L(1, 1)), "####.00"))
                    Else
                        L(3, 1) = 0
                    End If
                    If L(2, 2) <> 0 Then    '救濟
                        L(3, 2) = Format(Trim(L(2, 2) / L(1, 2)), "####.00")
                    Else
                        L(3, 2) = 0
                    End If
                    If L(2, 3) <> 0 Then    '其他
                        L(3, 3) = Format(Trim(L(2, 3) / L(1, 3)), "####.00")
                    Else
                        L(3, 3) = 0
                    End If
                    If L(2, 4) <> 0 Then    '空白  暫時無用
                        L(3, 4) = 0
                    Else
                        L(3, 4) = 0
                    End If
                '介於
                    If L(2, 5) <> 0 Then
                        L(3, 5) = Format(Trim(L(2, 5) / L(1, 5)), "####.00")  '爭議
                    Else
                        L(3, 5) = 0
                    End If
                    If L(2, 6) <> 0 Then
                        L(3, 6) = Format(Trim(L(2, 6) / L(1, 6)), "####.00")  '救濟
                    Else
                        L(3, 6) = 0
                    End If
                    If L(2, 7) <> 0 Then
                        L(3, 7) = Format(Trim(L(2, 7) / L(1, 7)), "####.00") '其他
                    Else
                        L(3, 7) = 0
                    End If
                    If L(2, 8) <> 0 Then
                        L(3, 8) = 0  '空白  暫時無用
                    Else
                        L(3, 8) = 0
                    End If
                '低標
                    If L(2, 9) <> 0 Then
                        L(3, 9) = Format(Trim(L(2, 9) / L(1, 9)), "####.00")  '爭議
                    Else
                        L(3, 9) = 0
                    End If
                    If L(2, 10) <> 0 Then
                        L(3, 10) = Format(Trim(L(2, 10) / L(1, 10)), "####.00") '救濟
                    Else
                        L(3, 10) = 0
                    End If
                    If L(2, 11) <> 0 Then
                        L(3, 11) = Format(Trim(L(2, 11) / L(1, 11)), "####.00") '其他
                    Else
                        L(3, 11) = 0
                    End If
                    If L(2, 12) <> 0 Then
                        L(3, 12) = 0 '空白  暫時無用
                    Else
                        L(3, 12) = 0
                    End If
                '標準價   爭議救濟 沒有標準價
                L(4, 1) = 0
                L(4, 2) = 0
                L(4, 3) = 0
                L(4, 4) = 0
                L(4, 5) = 0
                L(4, 6) = 0
                L(4, 7) = 0
                L(4, 8) = 0
                L(4, 9) = 0
                L(4, 10) = 0
                L(4, 11) = 0
                L(4, 12) = 0
                '差距
                '高標
                     L(5, 1) = Format(IIf(L(3, 1) = 0, 0, L(3, 1) - L(4, 1)), "####.00") '爭議
                     L(5, 2) = Format(IIf(L(3, 2) = 0, 0, L(3, 2) - L(4, 2)), "####.00") '救濟
                     L(5, 3) = Format(IIf(L(3, 3) = 0, 0, L(3, 3) - L(4, 3)), "####.00") '其他
                     L(5, 4) = 0 '空白  暫時無用
                '介於
                     L(5, 5) = Format(IIf(L(3, 5) = 0, 0, L(3, 5) - L(4, 5)), "####.00") '爭議
                     L(5, 6) = Format(IIf(L(3, 6) = 0, 0, L(3, 6) - L(4, 6)), "####.00") '救濟
                     L(5, 7) = Format(IIf(L(3, 7) = 0, 0, L(3, 7) - L(4, 7)), "####.00") '其他
                     L(5, 8) = 0 '空白  暫時無用
                '低標
                     L(5, 9) = Format(IIf(L(3, 9) = 0, 0, L(3, 9) - L(4, 9)), "####.00") '爭議
                     L(5, 10) = Format(IIf(L(3, 10) = 0, 0, L(3, 10) - L(4, 10)), "####.00") '救濟
                     L(5, 11) = Format(IIf(L(3, 11) = 0, 0, L(3, 11) - L(4, 11)), "####.00") '其他
                     L(5, 12) = 0 '空白  暫時無用
                '高標差    其他不算
                Hightmp = (L(1, 1) * L(5, 1)) + (L(1, 2) * L(5, 2)) '+ (L(1, 3) * L(5, 3))
                '介於差    其他不算
                Middletmp = (L(1, 5) * L(5, 5)) + (L(1, 6) * L(5, 6)) '+ (L(1, 7) * L(5, 7))
                '低標差    其他不算
                Lowtmp = (L(1, 9) * L(5, 9)) + (L(1, 10) * L(5, 10))   '+ (L(1, 11) * L(5, 11))
                '總和
                Sumtmp = Hightmp + Middletmp + Lowtmp
                '高標，介於，低標合計件       其他不算
                SumCount = L(1, 1) + L(1, 2) + L(1, 5) + L(1, 6) + L(1, 9) + L(1, 10)
                '平均差
                If SumCount <> 0 Then
                    Avgtmp = Format(IIf(SumCount <> 0, Sumtmp / SumCount, 0), "####.00")
                Else
                    Avgtmp = 0
                End If
                '存入暫存檔
                strSql = "insert into r090607_3 (r101001,r101002,r101003,r101004,r101005,r101006,r101007,r101008,id) values ('" & CheckStr(.Fields("r100001").Value) & "','" & CheckStr(.Fields("r100015").Value) & "'," & _
                         L(1, 1) + L(1, 2) + L(1, 3) & "," & L(1, 5) + L(1, 6) + L(1, 7) & "," & L(1, 9) + L(1, 10) + L(1, 11) & "," & Middletmp & "," & Sumtmp & "," & Avgtmp & ",'" & strUserNum & "')"
                cnnConnection.Execute strSql
                .MoveNext
            Loop
        End If
    End With
    CheckOC
End If
CheckOC
PrintData
End Sub

Sub PrintData()
'算月份  91.08.07  nick
'G = DateDiff("d", ChangeWStringToWDateString(ChangeTStringToWString(txt1(3) & "01")), ChangeWStringToWDateString(ChangeTStringToWString(txt1(4) & "31")))
'If G > 30 Then
'同年
allG = 0
'If Mid(ChangeTStringToWString(txt1(3) & "01"), 1, 4) = Mid(ChangeTStringToWString(txt1(4) & "31"), 1, 4) Then
'    If Mid(ChangeTStringToWString(txt1(3) & "01"), 5, 2) = Mid(ChangeTStringToWString(txt1(4) & "31"), 5, 2) Then
'    '同月
'        tmpnickG = 1
'    Else
'    '不同月
'        tmpnickG = DateDiff("m", ChangeWStringToWDateString(ChangeTStringToWString(txt1(3) & "01")), ChangeWStringToWDateString(ChangeTStringToWString(txt1(4) & "31"))) + 1
'    End If
'Else
'    '不同年
'    tmpnickG = DateDiff("m", ChangeWStringToWDateString(ChangeTStringToWString(txt1(3) & "01")), ChangeWStringToWDateString(ChangeTStringToWString(txt1(4) & "31"))) + 1
'End If
''  ****** end
'新申請案  爭議救濟案
strSql = "SELECT nvl(a0902,a0903),SUBStr(R101002,1,6),SUM(R101003),SUM(R101004),SUM(R101005),SUM(R101006),SUM(R101007),SUM(R101008),R101001 FROM R090607_3,acc090 WHERE R101001=a0901(+) and ID='" & strUserNum & "' GROUP BY R101001,SUBStr(R101002,1,6),nvl(a0902,a0903) ORDER BY R101001,SUBSTR(R101002,1,6),nvl(a0902,a0903) "
CheckOC
SavDay1 = ""
SavDay2 = "     "
Page = 1
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog ("") 'Add By Sindy 2010/12/20
        .MoveFirst
        SavDay1 = CheckStr(.Fields(8))
        PrintTitle
        tmpnickG = 0
        Do While .EOF = False
            For i = 0 To 8
                strTemp(i) = CheckStr(.Fields(i))
                If Len(strTemp(i)) = 0 And i > 1 Then
                    strTemp(i) = "0"
                End If
            Next i
            If strTemp(8) <> SavDay1 Then
                ShowLine
                PrintEnd1
                iPrint = iPrint + 600
                If iPrint >= 14000 Then
                    Page = Page + 1
                    Printer.NewPage
                    PrintTitle
                End If
                SavDay1 = strTemp(8)
                tmpnickG = 0
            End If
            tmpnickG = tmpnickG + 1
            If SavDay2 = strTemp(0) Then
                strTemp(0) = ""
            Else
                SavDay2 = strTemp(0)
            End If
            strTemp(0) = StrToStr(strTemp(0), 4)
            strTemp(1) = str(Val(Mid(strTemp(1), 1, 4)) - 1911) & "/" & Mid(strTemp(1), 5, 2)
            PrintDatil
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/20
        ShowNoData
        Exit Sub
    End If
End With
ShowLine
PrintEnd1
ShowLine2
PrintEnd2
ShowLine2
Printer.EndDoc
ShowPrintOk
Printer.Orientation = 2
End Sub

Sub PrintTitle() '列印抬頭
GetPleft
iPrint = 0
Printer.Orientation = 1
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 3000
Printer.CurrentY = iPrint
Printer.Print "智權人員收文高低標統計表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 3500
Printer.CurrentY = iPrint
'Printer.Print "收文年月：" & Mid(txt1(3), 1, 2) & "/" & Mid(txt1(3), 3, 2) & "－" & Mid(txt1(4), 1, 2) & "/" & Mid(txt1(4), 3, 2)
Printer.Print "收文年月：" & (txt1(3) \ 100) & "/" & Right(txt1(3), 2) & "－" & (txt1(4) \ 100) & "/" & Right(txt1(4), 2)
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 8500
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
If Val(txt1(8)) = 1 Then
    Printer.Print "統計性質：新申請案"
Else
    Printer.Print "統計性質：爭議/救濟案"
End If
Printer.CurrentX = 8500
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(11000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(3) + (Printer.TextWidth("介於") / 2) - (Printer.TextWidth("件數部分") / 2)
Printer.CurrentY = iPrint
Printer.Print "件數部分"
Printer.CurrentX = PLeft(6) + (Printer.TextWidth("綜合") / 2) - (Printer.TextWidth("點數分析") / 2)
Printer.CurrentY = iPrint
Printer.Print "點數分析"
iPrint = iPrint + 300
Printer.Line (PLeft(2), iPrint + 150)-(11000, iPrint + 150)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "業務區別"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "收文年月"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "高標"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "介於"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "低標"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "介於差"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "綜合"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "平均差距"
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(11000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
End Sub

Sub PrintDatil() '列印資料

Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
For i = 2 To 4
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "####0"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0")
Next i
For i = 5 To 7
    Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(strTemp(i), "####0.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(strTemp(i), "####0.00")
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
'定陣列
Erase PLeft
PLeft(0) = 0
PLeft(1) = 1200
PLeft(2) = 3000
PLeft(3) = 4330
PLeft(4) = 5660
PLeft(5) = 6990
PLeft(6) = 8320
PLeft(7) = 9650
End Sub

Sub ShowLine()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Line (0, iPrint + 150)-(11000, iPrint + 150)
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
End Sub

Sub ShowLine2()
Printer.CurrentX = 0
Printer.CurrentY = iPrint
For i = 120 To 125
    Printer.Line (0, iPrint + i)-(11000, iPrint + i)
    Printer.Line (0, iPrint + 50 + i)-(11000, iPrint + 50 + i)
Next i
iPrint = iPrint + 300
If iPrint >= 14000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If
End Sub

Sub PrintEnd1()
Dim tmpSubTotle As String
Dim tmpAvgTotle As String
allG = allG + tmpnickG
'列印結尾
'新申請案  爭議救濟案
strSql = "SELECT '小  計','',SUM(R101003),SUM(R101004),SUM(R101005),SUM(R101006),SUM(R101007),SUM(R101008) FROM R090607_3 WHERE ID='" & strUserNum & "' AND R101001='" & SavDay1 & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 7
                StrTemp7(i) = CheckStr(.Fields(i))
                If Len(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 4
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "####0")
            Next i
            'For i = 5 To 6
            '    Printer.CurrentX = Pleft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0.00"))
            '    Printer.CurrentY = iPrint
            '    Printer.Print Format(StrTemp7(i), "####0.00")
            'Next i
            'If Val(StrTemp7(3)) <> 0 Then
            '    tmpSubTotle = Format(StrTemp7(6) / StrTemp7(3), "####0.00")
            'Else
            '    tmpSubTotle = 0
            'End If
            'Printer.CurrentX = Pleft(7) + 500 - Printer.TextWidth(Format(tmpSubTotle, "####0.00"))
            'Printer.CurrentY = iPrint
            'Printer.Print Format(tmpSubTotle, "####0.00")
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            ShowLine
            StrTemp7(0) = "平  均"
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 4
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format((Val(StrTemp7(i)) / tmpnickG), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format((Val(StrTemp7(i)) / tmpnickG), "####0.00")
            Next i
            For i = 5 To 6
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format((Val(StrTemp7(i)) / tmpnickG), "####0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format((Val(StrTemp7(i)) / tmpnickG), "####0.00")
            Next i
            If ((Val(StrTemp7(3)) / tmpnickG) + (Val(StrTemp7(2)) / tmpnickG) + (Val(StrTemp7(4)) / tmpnickG)) <> 0 Then
                tmpAvgTotle = Format(((Val(StrTemp7(6)) / tmpnickG) / ((Val(StrTemp7(3)) / tmpnickG) + (Val(StrTemp7(2)) / tmpnickG) + (Val(StrTemp7(4)) / tmpnickG))), "####0.00")
            Else
                tmpAvgTotle = "0"
            End If
            Printer.CurrentX = PLeft(7) + 500 - Printer.TextWidth(Format(tmpAvgTotle, "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(tmpAvgTotle, "####0.00")
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub PrintEnd2()
'列印結尾
Dim tmpSubTotle As String
Dim tmpAvgTotle As String

strSql = "SELECT '總  計','',SUM(R101003),SUM(R101004),SUM(R101005),SUM(R101006),SUM(R101007),SUM(R101008) FROM R090607_3 WHERE ID='" & strUserNum & "' "
CheckOC2
With adoRecordset1
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 7
                StrTemp7(i) = CheckStr(.Fields(i))
                If Len(StrTemp7(i)) = 0 And i > 1 Then
                    StrTemp7(i) = "0"
                End If
            Next i
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 4
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format(StrTemp7(i), "####0")
            Next i
            'For i = 5 To 7
            '    Printer.CurrentX = Pleft(i) + 500 - Printer.TextWidth(Format(StrTemp7(i), "####0.00"))
            '    Printer.CurrentY = iPrint
            '    Printer.Print Format(StrTemp7(i), "####0.00")
            'Next i
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            ShowLine2
            StrTemp7(0) = "總平均"
            Printer.CurrentX = PLeft(0)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(0)
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iPrint
            Printer.Print StrTemp7(1)
            For i = 2 To 4
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format((Val(StrTemp7(i)) / tmpnickG), "####0"))
                Printer.CurrentY = iPrint
                Printer.Print Format((Val(StrTemp7(i)) / tmpnickG), "####0.00")
            Next i
            For i = 5 To 6
                Printer.CurrentX = PLeft(i) + 500 - Printer.TextWidth(Format((Val(StrTemp7(i)) / tmpnickG), "####0.00"))
                Printer.CurrentY = iPrint
                Printer.Print Format((Val(StrTemp7(i)) / tmpnickG), "####0.00")
            Next i
            If ((Val(StrTemp7(3)) / tmpnickG) + (Val(StrTemp7(2)) / tmpnickG) + (Val(StrTemp7(4)) / tmpnickG)) <> 0 Then
                tmpAvgTotle = Format(((Val(StrTemp7(6)) / tmpnickG) / ((Val(StrTemp7(3)) / tmpnickG) + (Val(StrTemp7(2)) / tmpnickG) + (Val(StrTemp7(4)) / tmpnickG))), "####0.00")
            Else
                tmpAvgTotle = "0"
            End If
            Printer.CurrentX = PLeft(7) + 500 - Printer.TextWidth(Format(tmpAvgTotle, "####0.00"))
            Printer.CurrentY = iPrint
            Printer.Print Format(tmpAvgTotle, "####0.00")
            iPrint = iPrint + 300
            If iPrint >= 14000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC2
End Sub

Sub Process2()
cnnConnection.Execute "DELETE FROM R090607_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090607_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
'系統類別
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 + " and CP01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) 'Add By Sindy 2010/12/20
End If
StrSQL6 = ""
'申請國家
If Len(txt1(1)) <> 0 Then
   strSQL1 = strSQL1 + " AND PA09='" & txt1(1) & "' "
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & lbl1(1) 'Add By Sindy 2010/12/20
End If
StrSQL6 = StrSQL6 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(3) & "01")) & " AND cp05<=" & Val(ChangeTStringToWString(txt1(4) & "31")) & " "
pub_QL05 = pub_QL05 & ";" & Label1(2) & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/12/20
'業務區
If Len(txt1(5)) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP12>='" & txt1(5) & "' "
End If
If Len(txt1(6)) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP12<='" & txt1(6) & "' "
End If
If Len(txt1(5)) <> 0 Or Len(txt1(6)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(3) & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/12/20
End If
'智權人員
If Len(txt1(7)) <> 0 Then
   StrSQL6 = StrSQL6 + " AND CP13='" & txt1(7) & "'  "
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(7) & lbl1(0) 'Add By Sindy 2010/12/20
End If
'Modify By Cheng 2003/07/07
'件數不含不計件資料, 點數不論是否計件都要含
'StrSQL6 = StrSQL6 + " and CP26 IS NULL  "
'2010/7/14 modify by sonia 與王副總再確認,不計件案件仍要計算件數和點數,但取消收文不列入故加入CP57的控制
'新申請案
'高標 : 點數>標準價
strSql = "select cp12,NVL(ST02,CP13)," & _
   "decode(cp10,'101',cp18,'104',decode(pa08,'1',cp18,0),0)," & _
   "decode(cp10,'102',cp18,'104',decode(pa08,'2',cp18,0),0)," & _
   "decode(cp10,'103',cp18,'105',cp18,'125',cp18,0)," & _
   "decode(cp10,'107',cp18,0)," & _
   "decode(cp10,'101',cp18,'102',cp18,'104',cp18,'105',cp18,'125',cp18,'107',cp18,'103',cp18,0)," & _
   "0,0,0,0,0,0,0,0,0,0,CP05,'" & strUserNum & "', CP26, CP09  from caseprogress,patent,staff,acc090 where " & _
   "CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and " & _
   "cp18>DECODE(cp33,NULL,0,CP33) and cp12=A0901(+) AND CP18>0 AND CP57 IS NULL " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,cp18,0,0,cp18,0,0,0,0,0,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and (cp10='102' or (cp10='104' and pa08='2'))  and cp18>cp33 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,cp18,0,cp18,0,0,0,0,0,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and CP10 IN ('103','105') and cp18>cp33 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,cp18,cp18,0,0,0,0,0,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and CP10='107' and cp18>cp33 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'介於 : 低價<=點數<=標準價
strSql = strSql + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,decode(cp10,'101',cp18,'104',decode(pa08,'1',cp18,0),0),decode(cp10,'102',cp18,'104',decode(pa08,'2',cp18,0),0),decode(cp10,'103',cp18,'105',cp18,'125',cp18,0),decode(cp10,'107',cp18,0),decode(cp10,'101',cp18,'102',cp18,'104',cp18,'105',cp18,'125',cp18,'107',cp18,'103',cp18,0),0,0,0,0,0,CP05,'" & strUserNum & "', CP26, CP09  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and cp18<=cp33 and cp18>=DECODE(cp34,NULL,0,CP34) and cp12=a0901(+) AND CP18>0 AND CP57 IS NULL " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,cp18,0,0,cp18,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and (cp10='102' or (cp10='104' and pa08='2'))  and cp18<=cp33 and cp18>=cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,cp18,0,cp18,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and CP10 IN ('103','105') and cp18<=cp33 and cp18>=cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,cp18,cp18,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and CP10='107' and cp18<=cp33 and cp18>=cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'低底 : 底價<點數
strSql = strSql + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,0,0,decode(cp10,'101',cp18,'104',decode(pa08,'1',cp18,0),0),decode(cp10,'102',cp18,'104',decode(pa08,'2',cp18,0),0),decode(cp10,'103',cp18,'105',cp18,'125',cp18,0),decode(cp10,'107',cp18,0),decode(cp10,'101',cp18,'102',cp18,'104',cp18,'105',cp18,'125',cp18,'107',cp18,'103',cp18,0),CP05,'" & strUserNum & "', CP26, CP09  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and cp18<cp34 and cp12=a0901(+) AND CP18>0 AND CP57 IS NULL " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,0,0,0,cp18,0,0,cp18,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and (cp10='102' or (cp10='104' and pa08='2'))  and cp18<cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,0,0,0,0,cp18,0,cp18,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and CP10 IN ('103','105') and cp18<cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,0,0,0,0,0,cp18,cp18,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and CP10='107' and cp18<cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
cnnConnection.Execute "insert into r090607_1 " & strSql
'CheckOC
'爭議救濟案
                strSql = "select cp12,NVL(ST02,CP13),decode(substr(cp10,1,1),'8',cp18,0),decode(substr(cp10,1,1),'5',cp18,0),decode(substr(cp10,1,1),'5',0,'8',0,cp18),cp18,0,0,0,0,0,0,0,0,CP05,'" & strUserNum & "', CP26 from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and cp18>DECODE(cp33,NULL,0,CP33) and cp12=A0901(+) AND CP18>0 AND CP57 IS NULL " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,cp18,0,cp18,0,0,0,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and SUBSTR(cp10,1,1)='5' and cp18>cp33 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,cp18,cp18,0,0,0,0,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and (SUBSTR(cp10,1,1)<>'8' AND SUBSTR(cp10,1,1)<>'5') and cp18>cp33 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
strSql = strSql + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,decode(substr(cp10,1,1),'8',cp18,0),decode(substr(cp10,1,1),'5',cp18,0),decode(substr(cp10,1,1),'5',0,'8',0,cp18),cp18,0,0,0,0,CP05,'" & strUserNum & "', CP26 from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and cp18<=cp33 and cp18>=DECODE(cp34,NULL,0,CP34) and cp12=a0901(+) AND CP18>0 AND CP57 IS NULL " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,cp18,0,cp18,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and SUBSTR(cp10,1,1)='5' and cp18<=cp33 and cp18>=cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,cp18,cp18,0,0,0,0,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and (SUBSTR(cp10,1,1)<>'8' AND SUBSTR(cp10,1,1)<>'5') and cp18<=cp33 and cp18>=cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
strSql = strSql + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,decode(substr(cp10,1,1),'8',cp18,0),decode(substr(cp10,1,1),'5',cp18,0),decode(substr(cp10,1,1),'5',0,'8',0,cp18),cp18,CP05,'" & strUserNum & "', CP26 from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and cp18<cp34 and cp12=a0901(+) AND CP18>0 AND CP57 IS NULL " & strSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,0,cp18,0,cp18,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and SUBSTR(cp10,1,1)='5' and cp18<cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
'strSQL = strSQL + " UNION all  SELECT cp12,NVL(ST02,CP13),0,0,0,0,0,0,0,0,0,0,cp18,cp18,CP05,'" & strUserNum & "'  from caseprogress,patent,staff,acc090 where CP01=pa01(+) and CP02=pa02(+) and CP03=pa03(+) and CP04=pa04(+) AND CP13=ST01(+) and (SUBSTR(cp10,1,1)<>'8' AND SUBSTR(cp10,1,1)<>'5') and cp18<cp34 and cp12=a0901(+) AND CP18>0 " & StrSQL1 + StrSQL6
cnnConnection.Execute "insert into r090607_2 " & strSql
End Sub

Sub Process()
Select Case Val(txt1(9))
Case 1           '螢幕
     pub_QL05 = pub_QL05 & ";" & Label1(6) & "1.螢幕" 'Add By Sindy 2010/12/20
     Select Case Val(txt1(8))
     Case 1 '新申請案
          pub_QL05 = pub_QL05 & ";" & Label1(5) & "1.新申請案" 'Add By Sindy 2010/12/20
          Select Case Val(txt1(10))
          Case 1 '各區
               pub_QL05 = pub_QL05 & ";" & Label1(7) & "1.各區" 'Add By Sindy 2010/12/20
               CheckOC2
               strSql = "select * from r090607_1 where id='" & strUserNum & "' "
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 Then
                  InsertQueryLog ("") 'Add By Sindy 2010/12/20
                  Me.Hide
                  frm090607_1.Show
               Else
                  InsertQueryLog (0) 'Add By Sindy 2010/12/20
                  ShowNoData
               End If
               CheckOC2
          Case 2 '全部
               pub_QL05 = pub_QL05 & ";" & Label1(7) & "2.合併" 'Add By Sindy 2010/12/20
               CheckOC2
               strSql = "select * from r090607_1 where id='" & strUserNum & "' "
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 Then
                  InsertQueryLog ("") 'Add By Sindy 2010/12/20
                  Me.Hide
                  frm090607_2.Show
               Else
                  InsertQueryLog (0) 'Add By Sindy 2010/12/20
                  ShowNoData
               End If
               CheckOC2
          Case Else
          End Select
     Case 2 '爭議救濟案
          pub_QL05 = pub_QL05 & ";" & Label1(5) & "2.爭議/救濟案" 'Add By Sindy 2010/12/20
          Select Case Val(txt1(10))
          Case 1 '各區
               pub_QL05 = pub_QL05 & ";" & Label1(7) & "1.各區" 'Add By Sindy 2010/12/20
               CheckOC2
               strSql = "select * from r090607_1 where id='" & strUserNum & "' "
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 Then
                  InsertQueryLog ("") 'Add By Sindy 2010/12/20
                  Me.Hide
                  frm090607_3.Show
               Else
                  InsertQueryLog (0) 'Add By Sindy 2010/12/20
                  ShowNoData
               End If
               CheckOC2
          Case 2 '全部
               pub_QL05 = pub_QL05 & ";" & Label1(7) & "2.合併" 'Add By Sindy 2010/12/20
               CheckOC2
               strSql = "select * from r090607_1 where id='" & strUserNum & "' "
               adoRecordset1.CursorLocation = adUseClient
               adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If adoRecordset1.RecordCount <> 0 Then
                  InsertQueryLog ("") 'Add By Sindy 2010/12/20
                  Me.Hide
                  frm090607_4.Show
               Else
                  InsertQueryLog (0) 'Add By Sindy 2010/12/20
                  ShowNoData
               End If
               CheckOC2
          Case Else
          End Select
     Case Else
     End Select
Case 2              '報表
     pub_QL05 = pub_QL05 & ";" & Label1(6) & "2.報表" 'Add By Sindy 2010/12/20
     Process1
Case Else
End Select
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = Systemkind_g
For i = 0 To 7
    StrTemp99(i) = ""
Next i
txt1(9) = "1"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090607 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub Txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdok(0).SetFocus
End If
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '系統類別
      'Add By Cheng 2002/01/07
      Me.txt1(Index).Text = GetAllSysKind(Me.txt1(Index))
     If Len(txt1(0)) <> 0 Then
        If InStr(1, Systemkind_g, txt1(0) & ",") <> 0 Then
            If txt1(0) <> "FCP" And txt1(0) <> "P" And txt1(0) <> "CFP" Then
                s = MsgBox("此表單不適用在 " & txt1(0) & " 的系統類別!!", , "USER 系統類別使用錯誤")
                txt1(0).SetFocus
                txt1(0).SelStart = 0
                txt1(0).SelLength = Len(txt1(0))
                Exit Sub
            End If
        Else
            s = MsgBox("USER " & strUserNum & " 沒有" & txt1(0) & " 的使用權限!!", , "USER 沒有權限")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    End If
Case 1
     lbl1(1).Caption = GetPrjNationNameHM(txt1(1))
          If Trim(txt1(Index)) <> "" Then
        If Trim(lbl1(1).Caption) = "" Then
            s = MsgBox("申請國家輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 3, 4
   If PUB_CheckKeyInYYMM(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 4 Then
    If RunNick(txt1(Index - 1), txt1(Index)) Then
        txt1(Index - 1).SetFocus
        txt1_GotFocus (Index - 1)
        Exit Sub
    End If
   End If
Case 6
   If RunNick(txt1(Index - 1), txt1(Index)) Then
       txt1(Index - 1).SetFocus
       txt1_GotFocus (Index - 1)
       Exit Sub
   End If
Case 7
     lbl1(0).Caption = GetPrjSalesNM(txt1(7))
     If Trim(txt1(Index)) <> "" Then
        If Trim(lbl1(0).Caption) = "" Then
            s = MsgBox("智權人員輸入錯誤！", , "錯誤！")
            txt1(Index).SetFocus
            txt1_GotFocus (Index)
            Exit Sub
        End If
     End If
Case 8
     Select Case Trim(txt1(8))
     Case "1", "2", ""
     Case Else
          s = MsgBox("查詢性質只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(8).SetFocus
          txt1(8).SelStart = 0
          txt1(8).SelLength = Len(txt1(8))
          Exit Sub
     End Select
Case 9
     Select Case Trim(txt1(9))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能 1 或 2 !!", , "USER 輸入錯誤")
          txt1(9).SetFocus
          txt1(9).SelStart = 0
          txt1(9).SelLength = Len(txt1(9))
          Exit Sub
     End Select
Case 10
     Select Case Trim(txt1(10))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示內容只能 1 或 2 !!", , "USER 輸入錯誤")
          txt1(10).SetFocus
          txt1(10).SelStart = 0
          txt1(10).SelLength = Len(txt1(10))
          Exit Sub
     End Select
Case Else
End Select
End Sub


