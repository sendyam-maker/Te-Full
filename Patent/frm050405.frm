VERSION 5.00
Begin VB.Form frm050405 
   BorderStyle     =   1  '單線固定
   Caption         =   "代理人/申請人新申請案排行榜"
   ClientHeight    =   3270
   ClientLeft      =   2505
   ClientTop       =   3060
   ClientWidth     =   3540
   ControlBox      =   0   'False
   LinkTopic       =   "Form24"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   3540
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   930
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2490
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   9
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   4
      Top             =   1140
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   930
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1155
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   924
      TabIndex        =   0
      Top             =   504
      Width           =   2535
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   930
      MaxLength       =   3
      TabIndex        =   1
      Top             =   816
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   924
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1485
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   924
      MaxLength       =   7
      TabIndex        =   6
      Top             =   1815
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   924
      MaxLength       =   1
      TabIndex        =   8
      Top             =   2160
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   924
      MaxLength       =   4
      TabIndex        =   10
      Top             =   2850
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2736
      TabIndex        =   12
      Top             =   48
      Width           =   756
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1944
      TabIndex        =   11
      Top             =   48
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   2244
      MaxLength       =   7
      TabIndex        =   7
      Top             =   1815
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2130
      MaxLength       =   3
      TabIndex        =   2
      Top             =   804
      Width           =   800
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "排行條件："
      Height          =   180
      Left            =   90
      TabIndex        =   24
      Top             =   2535
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.CF件數 2.FC件數)"
      Height          =   180
      Left            =   1260
      TabIndex        =   23
      Top             =   2535
      Width           =   1575
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "國籍："
      Height          =   180
      Left            =   90
      TabIndex        =   22
      Top             =   1200
      Width           =   540
   End
   Begin VB.Line Line3 
      X1              =   1800
      X2              =   2040
      Y1              =   1275
      Y2              =   1275
   End
   Begin VB.Line Line2 
      X1              =   1890
      X2              =   2130
      Y1              =   1935
      Y2              =   1935
   End
   Begin VB.Line Line1 
      X1              =   1800
      X2              =   2040
      Y1              =   930
      Y2              =   930
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(空白表全部)"
      Height          =   180
      Left            =   1800
      TabIndex        =   21
      Top             =   2880
      Width           =   1020
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印名次："
      Height          =   180
      Left            =   90
      TabIndex        =   20
      Top             =   2880
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1.代理人 2.申請人)"
      Height          =   180
      Left            =   1260
      TabIndex        =   19
      Top             =   2205
      Width           =   1515
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印對象："
      Height          =   180
      Left            =   90
      TabIndex        =   18
      Top             =   2205
      Width           =   900
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "日  期："
      Height          =   180
      Left            =   90
      TabIndex        =   17
      Top             =   1845
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "(1. 收文 2.發文)"
      Height          =   180
      Left            =   1260
      TabIndex        =   16
      Top             =   1515
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "列印別："
      Height          =   180
      Left            =   90
      TabIndex        =   15
      Top             =   1515
      Width           =   720
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "申請國家："
      Height          =   180
      Left            =   90
      TabIndex        =   14
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "系統類別："
      Height          =   180
      Left            =   90
      TabIndex        =   13
      Top             =   540
      Width           =   900
   End
End
Attribute VB_Name = "frm050405"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/29 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, strTemp1 As Variant, strTemp2 As Variant
Dim iPrint As Integer, Page As Integer, k As Integer, PLeft(0 To 5) As Integer, strTemp(0 To 7) As String, StrTest As String
Dim StrTest2 As String, StrTemp5(0 To 1) As String, StrTest99 As String, iY As Integer, strSQL2 As String

Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0
     Printer.Orientation = 2
     DoEvents
     If Len(txt1(0)) = 0 And txt1(6) <> "1" Then
         s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
        'Modify By Cheng 2003/11/11
        '取消限制
'         If Len(txt1(2)) = 0 Then
'            'Modify By cheng 2002/02/19
''            s = MsgBox("申請國家或國籍區間不可空白!!", , "USER 輸入錯誤")
'            s = MsgBox("申請國家區間不可空白!!", , "USER 輸入錯誤")
'            txt1(1).SetFocus
'            txt1_GotFocus (1)
'            Exit Sub
'         Else
            If Len(txt1(3)) = 0 Then
                s = MsgBox("列印別不可空白!!", , "USER 輸入錯誤")
                txt1(3).SetFocus
                Exit Sub
            Else
                'Add By Cheng 2002/03/19
               If PUB_CheckKeyInDate(Me.txt1(4)) = -1 Then
                  Me.txt1(4).SetFocus
                  txt1_GotFocus 4
                  Exit Sub
               End If
               If PUB_CheckKeyInDate(Me.txt1(5)) = -1 Then
                  Me.txt1(5).SetFocus
                  txt1_GotFocus 5
                  Exit Sub
               End If
                
                If Len(txt1(5)) = 0 Then
                    s = MsgBox("日期區間不可空白!!", , "USER 輸入錯誤")
                    txt1(4).SetFocus
                    txt1_GotFocus (4)
                    Exit Sub
                Else
                    If Len(txt1(6)) = 0 Then
                        'Modify By Cheng 2002/02/19
'                        s = MsgBox("排行條件不可空白!!", , "USER 輸入錯誤")
                        s = MsgBox("列印對象不可空白!!", , "USER 輸入錯誤")
                        txt1(6).SetFocus
                        Exit Sub
                    Else
                    'Modify By Cheng 2003/11/11
                    '取消限制
'                     'Add By Cheng 2002/02/19
'                     If Len(txt1(9)) = 0 Then
'                        s = MsgBox("國籍區間不可空白!!", , "USER 輸入錯誤")
'                        txt1(8).SetFocus
'                        txt1_GotFocus (8)
'                        Exit Sub
'                     End If
                    If Len(txt1(10)) = 0 And Me.txt1(6).Text = "1" Then
                        'Modify By Cheng 2002/02/19
                        s = MsgBox("排行條件不可空白!!", , "USER 輸入錯誤")
                        txt1(10).SetFocus
                        Exit Sub
                     End If
                                                
                        Screen.MousePointer = vbHourglass
                        Me.Enabled = False
                        StrTest = StrTest2
                        'strTemp1 = Split(Replace(UCase(StrTest), ",,", ""), ",")
                        'strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
                        'For i = 0 To UBound(strTemp1)
                        '    S = 0
                        '    For j = 0 To UBound(strTemp2)
                        '        If strTemp2(j) = strTemp1(i) Then
                        '            S = 1
                        '            Exit For
                        '        End If
                        '    Next j
                        '    If S = 0 Then
                        '        StrTest = Replace(StrTest, "," & strTemp1(i) & ",", "")
                        '    End If
                        'Next i
                        'Me.Tag = strTemp2(0)
                        ClearQueryLog (Me.Name) 'Add By Sindy 2010/10/19 清除查詢印表記錄檔欄位
                        If Len(txt1(10)) <> 0 Then
                           pub_QL05 = pub_QL05 & ";" & Label12 & txt1(10) & Label11 'Add By Sindy 2010/10/19
                        End If
                        If Len(txt1(7)) <> 0 Then
                           pub_QL05 = pub_QL05 & ";" & Label8 & txt1(7) 'Add By Sindy 2010/10/19
                        End If
                        If txt1(6) = "1" Then '列印對象--代理人
                            pub_QL05 = pub_QL05 & ";" & Label6 & "代理人" 'Add By Sindy 2010/10/19
                            Process
                        Else '列印對象--申請人
                            pub_QL05 = pub_QL05 & ";" & Label6 & "申請人" 'Add By Sindy 2010/10/19
                            Process1
                        End If
                        Me.Enabled = True
                        Screen.MousePointer = vbDefault
                    End If
                End If
            End If
'         End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

'代理人
Private Sub Process()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050405_1 where id='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
Select Case intPCaseKind
Case 1
'      strSQL1 = strSQL1 + " AND CP01 IN ('CFP','FCP','P') "
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/10/19
      'Modify By Cheng 2002/02/19
'      strSQL1 = strSQL1 + " AND ((PA09>='" & txt1(1) & "' AND PA09<='" & txt1(2) & "') OR (F1.FA10>='" & txt1(1) & "' AND F1.FA10<='" & txt1(2) & "z') OR (F2.FA10>='" & txt1(1) & "' AND F2.FA10<='" & txt1(2) & "z')) "
        'Modify By Cheng 2003/11/11
'      strSQL1 = strSQL1 + " AND (PA09>='" & txt1(1) & "' AND PA09<='" & txt1(2) & "') AND ((F1.FA10>='" & txt1(8) & "' AND F1.FA10<='" & txt1(9) & "z') OR (F2.FA10>='" & txt1(8) & "' AND F2.FA10<='" & txt1(9) & "z')) "
        If Trim(Me.txt1(1).Text) <> "" Then
            strSQL1 = strSQL1 & " And PA09>='" & Me.txt1(1).Text & "' "
        Else
            strSQL1 = strSQL1 & " And PA09=PA09 "
        End If
        If Trim(Me.txt1(2).Text) <> "" Then
            strSQL1 = strSQL1 & " And PA09<='" & Me.txt1(2).Text & "' "
        End If
        If Trim(Me.txt1(1).Text) <> "" Or Trim(Me.txt1(2).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
        End If
        If Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL1 = strSQL1 & " AND ((F1.FA10>='" & txt1(8) & "' AND F1.FA10<='" & txt1(9) & "z') OR (F2.FA10>='" & txt1(8) & "' AND F2.FA10<='" & txt1(9) & "z')) "
        ElseIf Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) = "" Then
            strSQL1 = strSQL1 & " AND (F1.FA10>='" & txt1(8) & "' OR F2.FA10>='" & txt1(8) & "') "
        ElseIf Trim(Me.txt1(8).Text) = "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL1 = strSQL1 & " AND (F1.FA10<='" & txt1(9) & "z' OR F2.FA10<='" & txt1(9) & "z') "
        Else
            strSQL1 = strSQL1 & " AND (F1.FA10=F1.FA10 OR F2.FA10=F2.FA10) "
        End If
        If Trim(Me.txt1(8).Text) <> "" Or Trim(Me.txt1(9).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label10 & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/10/19
        End If
Case 2
'      strSQL2 = strSQL2 + " AND CP01 IN ('T','TF','FCT','CFT') "
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/10/19
      'Modify By Cheng 2002/02/19
'      strSQL2 = strSQL2 + " AND ((TM10>='" & txt1(1) & "' AND TM10<='" & txt1(2) & "') OR (F1.FA10>='" & txt1(1) & "' AND F1.FA10<='" & txt1(2) & "z') OR (F2.FA10>='" & txt1(1) & "' AND F2.FA10<='" & txt1(2) & "z')) "
        'Modify By Cheng 2003/11/11
'      strSQL2 = strSQL2 + " AND (TM10>='" & txt1(1) & "' AND TM10<='" & txt1(2) & "') AND ((F1.FA10>='" & txt1(8) & "' AND F1.FA10<='" & txt1(9) & "z') OR (F2.FA10>='" & txt1(8) & "' AND F2.FA10<='" & txt1(9) & "z')) "
        If Trim(Me.txt1(1).Text) <> "" Then
            strSQL2 = strSQL2 & " And TM10>='" & Me.txt1(1).Text & "' "
        Else
            strSQL2 = strSQL2 & " And TM10=TM10 "
        End If
        If Trim(Me.txt1(2).Text) <> "" Then
            strSQL2 = strSQL2 & " And TM10<='" & Me.txt1(2).Text & "' "
        End If
        If Trim(Me.txt1(1).Text) <> "" Or Trim(Me.txt1(2).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
        End If
        If Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL2 = strSQL2 & " AND ((F1.FA10>='" & txt1(8) & "' AND F1.FA10<='" & txt1(9) & "z') OR (F2.FA10>='" & txt1(8) & "' AND F2.FA10<='" & txt1(9) & "z')) "
        ElseIf Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) = "" Then
            strSQL2 = strSQL2 & " AND (F1.FA10>='" & txt1(8) & "' OR F2.FA10>='" & txt1(8) & "') "
        ElseIf Trim(Me.txt1(8).Text) = "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL2 = strSQL2 & " AND (F1.FA10<='" & txt1(9) & "z' OR F2.FA10<='" & txt1(9) & "z') "
        Else
            strSQL2 = strSQL2 & " AND (F1.FA10=F1.FA10 OR F2.FA10=F2.FA10) "
        End If
        If Trim(Me.txt1(8).Text) <> "" Or Trim(Me.txt1(9).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label10 & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/10/19
        End If
Case Else
End Select
If Trim(txt1(3)) = "1" Then
   If Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(4))) & " "
      strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(4))) & " "
   End If
   If Len(Trim(txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(5))) & " "
      strSQL2 = strSQL2 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(5))) & " "
   End If
   If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";收文" & Label5 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/19
   End If
Else
   If Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(4))) & " "
      strSQL2 = strSQL2 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(4))) & " "
   End If
   If Len(Trim(txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(5))) & " "
      strSQL2 = strSQL2 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(5))) & " "
   End If
   If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";發文" & Label5 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/19
   End If
End If
'Modify By Cheng 2003/10/03
'Begin
'strSQL1 = strSQL1 + " AND (CP10='101' OR CP10='102' OR CP10='103' OR CP10='104' OR CP10='105') "
'2008/3/12 MODIFY BY SONIA 加CP09<'B'否則EPC指定國家子案也會計入
'strSQL1 = strSQL1 + " AND (CP10='101' OR CP10='102' OR CP10='103' OR CP10='104' OR CP10='105' Or CP10='109' Or CP10='110' Or CP10='112' Or CP10='113' Or CP10='114' Or CP10='115') "
strSQL1 = strSQL1 + " AND CP09<'B' AND (CP10='101' OR CP10='102' OR CP10='103' OR CP10='104' OR CP10='105' OR CP10='125' Or CP10='109' Or CP10='110' Or CP10='112' Or CP10='113' Or CP10='114' Or CP10='115') "
'End
'2008/3/12 MODIFY BY SONIA 加CP09<'B'否則EPC指定國家子案也會計入
'strSQL2 = strSQL2 + " AND (CP10='101') "
strSQL2 = strSQL2 + " AND CP09<'B' AND (CP10='101') "
CheckOC
Select Case intPCaseKind
Case 1
      strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,F2.FA10,CP44,PA75,F1.FA10,PA09 FROM FAGENT F1,FAGENT F2,PATENT,CASEPROGRESS " & _
         " WHERE CP01=pa01(+) AND CP02=pa02(+) AND CP03=pa03(+) AND CP04=pa04(+) AND CP57 IS NULL AND " & SQLNewFag("PA75", "F1.FA") & " AND " & SQLNewFag("CP44", "F2.FA") & strSQL1
Case 2
      strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,F2.FA10,CP44,TM44,F1.FA10,TM10 FROM FAGENT F1,FAGENT F2,TRADEMARK,CASEPROGRESS " & _
         " WHERE CP01=tm01(+) AND CP02=tm02(+) AND cP03=tm03(+) AND CP04=tm04(+) AND CP57 IS NULL AND " & SQLNewFag("TM44", "F1.FA") & " AND " & SQLNewFag("CP44", "F2.FA") & strSQL2
Case Else
End Select
strSql = strSql + " ORDER BY 2 "
'891024 改

k = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 5
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        If Val(Mid(strTemp(5), 1, 3)) = 0 Then            '國內
              If Len(Trim(strTemp(2))) = 0 And Len(Trim(strTemp(3))) = 0 Then
                  strTemp(0) = GetPrjNationName(strTemp(5))
                  strTemp(1) = ""
                  strTemp(2) = ""
                  strTemp(3) = "1"
                  strTemp(4) = ""
              Else
                  If Len(Trim(strTemp(2))) <> 0 Then
                     strTemp(0) = IIf(Trim(GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(2))))) = "", GetPrjNationNumber(GetNewFagent(strTemp(2))), GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(2)))))
                        'Modify By Cheng 2003/10/06
                        '若代理人國籍為台灣或大陸, 則代理人名稱中-->英-->日, 其他則英-->中-->日
                        'Begin
                        If strTemp(1) < "010" Or strTemp(1) = "020" Then
                            strTemp(1) = IIf(Trim(GetPrjName1(GetNewFagent(strTemp(2)))) = "", strTemp(2), GetPrjName1(GetNewFagent(strTemp(2))))
                        Else
                            strTemp(1) = IIf(Trim(GetPrjName2(GetNewFagent(strTemp(2)))) = "", strTemp(2), GetPrjName2(GetNewFagent(strTemp(2))))
                        End If
                        'End
                     strTemp(2) = GetNewFagent(strTemp(2))
                     strTemp(3) = ""
                     strTemp(4) = "1"
                  Else
                     strTemp(0) = IIf(Trim(GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(3))))) = "", GetPrjNationNumber(GetNewFagent(strTemp(3))), GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(3)))))
                        'Modify By Cheng 2003/10/06
                        '若代理人國籍為台灣或大陸, 則代理人名稱中-->英-->日, 其他則英-->中-->日
                        'Begin
                        If strTemp(4) < "010" Or strTemp(4) = "020" Then
                            strTemp(1) = IIf(Trim(GetPrjName1(GetNewFagent(strTemp(3)))) = "", strTemp(3), GetPrjName1(GetNewFagent(strTemp(3))))
                        Else
                            strTemp(1) = IIf(Trim(GetPrjName2(GetNewFagent(strTemp(3)))) = "", strTemp(3), GetPrjName2(GetNewFagent(strTemp(3))))
                        End If
                     strTemp(2) = GetNewFagent(strTemp(3))
                     strTemp(3) = ""
                     strTemp(4) = "1"
                  End If
              End If
          Else      '國外
              If Len(Trim(strTemp(2))) = 0 And Len(Trim(strTemp(3))) = 0 Then
                  strTemp(0) = GetPrjNationName(strTemp(5))
                  strTemp(1) = ""
                  strTemp(2) = ""
                  strTemp(3) = "1"
                  strTemp(4) = ""
              Else
                  If Len(Trim(strTemp(2))) <> 0 Then
                      strTemp(0) = IIf(Trim(GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(2))))) = "", GetPrjNationNumber(GetNewFagent(strTemp(2))), GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(2)))))
                        'Modify By Cheng 2003/10/06
                        '若代理人國籍為台灣或大陸, 則代理人名稱中-->英-->日, 其他則英-->中-->日
                        'Begin
                        If strTemp(1) < "010" Or strTemp(1) = "020" Then
                            strTemp(1) = IIf(Trim(GetPrjName1(GetNewFagent(strTemp(2)))) = "", strTemp(2), GetPrjName1(GetNewFagent(strTemp(2))))
                        Else
                            strTemp(1) = IIf(Trim(GetPrjName2(GetNewFagent(strTemp(2)))) = "", strTemp(2), GetPrjName2(GetNewFagent(strTemp(2))))
                        End If
                      strTemp(2) = GetNewFagent(strTemp(2))
                      strTemp(3) = "1"
                      strTemp(4) = ""
                  Else
                      strTemp(0) = IIf(Trim(GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(3))))) = "", GetPrjNationNumber(GetNewFagent(strTemp(3))), GetPrjNationName(GetPrjNationNumber(GetNewFagent(strTemp(3)))))
                        'Modify By Cheng 2003/10/06
                        '若代理人國籍為台灣或大陸, 則代理人名稱中-->英-->日, 其他則英-->中-->日
                        'Begin
                        If strTemp(4) < "010" Or strTemp(4) = "020" Then
                            strTemp(1) = IIf(Trim(GetPrjName1(GetNewFagent(strTemp(3)))) = "", strTemp(3), GetPrjName1(GetNewFagent(strTemp(3))))
                        Else
                            strTemp(1) = IIf(Trim(GetPrjName2(GetNewFagent(strTemp(3)))) = "", strTemp(3), GetPrjName2(GetNewFagent(strTemp(3))))
                        End If
                      strTemp(2) = GetNewFagent(strTemp(3))
                      strTemp(3) = "1"
                      strTemp(4) = ""
                  End If
              End If
          End If
         strSql = "INSERT INTO R050405_1 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & ",'" & strUserNum & "') "
         cnnConnection.Execute strSql
        DoEvents
        adoRecordset.MoveNext
    Loop
Else
    InsertQueryLog (0) 'Add By Sindy 2010/10/19
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
strSql = "SELECT R019001,R019002,R019003,SUM(R019004),SUM(R019005) FROM R050405_1 WHERE ID='" & strUserNum & "' GROUP BY R019001,R019002,R019003 "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/10/19
    adoRecordset.MoveFirst
    cnnConnection.Execute "DELETE FROM R050405_1 WHERE ID='" & strUserNum & "' "
    Do While adoRecordset.EOF = False
        For i = 0 To 4
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        strSql = "INSERT INTO R050405_1 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & "," & Val(strTemp(4)) & ",'" & strUserNum & "') "
        cnnConnection.Execute strSql
        adoRecordset.MoveNext
    Loop
Else
    InsertQueryLog (0)  'Add By Sindy 2010/10/19
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
PrintData

Screen.MousePointer = vbDefault
End Sub

Private Sub PrintData()
Select Case intPWhere
'modify by sonia 91.1.25
'Case 國內
'    strSQL = "SELECT * FROM R050405_1 WHERE ID='" & strUserNum & "' AND R019003 IS NOT NULL ORDER BY R019004 DESC,r019005 desc,R019003 ASC "
'Case Else
'    strSQL = "SELECT * FROM R050405_1 WHERE ID='" & strUserNum & "' AND R019003 IS NOT NULL ORDER BY R019005 DESC,r019004 desc,R019003 ASC "
'End Select
Case 國外_FC
      'Modify By Cheng 2002/02/19
'    strSQL = "SELECT * FROM R050405_1 WHERE ID='" & strUserNum & "' AND R019003 IS NOT NULL ORDER BY R019005 DESC,r019004 desc,R019003 ASC "
    strSql = "SELECT * FROM R050405_1 WHERE ID='" & strUserNum & "' AND R019003 IS NOT NULL ORDER BY " & IIf(Me.txt1(10).Text = "1", "R019004 DESC,r019005 desc", "R019005 DESC,r019004 desc") & ",R019003 ASC "
Case Else
      'Modify By Cheng 2002/02/19
'    strSQL = "SELECT * FROM R050405_1 WHERE ID='" & strUserNum & "' AND R019003 IS NOT NULL ORDER BY R019004 DESC,r019005 desc,R019003 ASC "
    strSql = "SELECT * FROM R050405_1 WHERE ID='" & strUserNum & "' AND R019003 IS NOT NULL ORDER BY " & IIf(Me.txt1(10).Text = "1", "R019004 DESC,r019005 desc", "R019005 DESC,r019004 desc") & ",R019003 ASC "
End Select

CheckOC
StrTest99 = "       "
iY = 0
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 1 To 5
                strTemp(i) = CheckStr(.Fields(i - 1))
            Next i
            strTemp(0) = strTemp(1)
'            If Mid(Me.Tag, 1, 2) = "FC" Then
            If Me.txt1(10).Text = "2" Then
                If StrTest99 <> strTemp(5) Then
                    StrTest99 = strTemp(5)
                    iY = iY + 1
                    strTemp(1) = Trim(str(iY))
                Else
                    strTemp(1) = Trim(str(iY))
                End If
            Else
                If StrTest99 <> strTemp(4) Then
                    StrTest99 = strTemp(4)
                    iY = iY + 1
                    strTemp(1) = Trim(str(iY))
                Else
                    strTemp(1) = Trim(str(iY))
                End If
            End If
            strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 12), vbUnicode)
            strTemp(2) = StrConv(MidB(StrConv(strTemp(2), vbFromUnicode), 1, 36), vbUnicode)
            If Len(txt1(7)) <> 0 Then
               If iY > Val(txt1(7)) Then
                  GoTo OutPrintDoc
               End If
            End If
            If iPrint > 10000 Then
                Page = Page + 1
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(250, "-")
                iPrint = iPrint + 300
                Printer.NewPage
                PrintTitle
            End If
            PrintDatil
            .MoveNext
        Loop
        'iY = .RecordCount
    End With
End If
CheckOC
OutPrintDoc:
strTemp(2) = "0"
strTemp(1) = "0"
strTemp(0) = "0"
'無代理人件數
strSql = "SELECT SUM(R019004), SUM(R019005) From R050405_1 WHERE ID='" & strUserNum & "' AND (R019002='' OR R019002 IS NULL)"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If Me.txt1(10).Text = "2" Then
        strTemp(1) = CheckStr(adoRecordset.Fields(1))
    Else
        strTemp(1) = CheckStr(adoRecordset.Fields(0))
    End If
End If
CheckOC
'收發文件數
strSql = "SELECT SUM(R019004),SUM(R019005) From R050405_1 WHERE ID='" & strUserNum & "' AND R019003 IS NOT NULL "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'    If Mid(Me.Tag, 1, 2) = "FC" Then
    If Me.txt1(10).Text = "2" Then
        strTemp(2) = CheckStr(adoRecordset.Fields(1))
    Else
        strTemp(2) = CheckStr(adoRecordset.Fields(0))
    End If
End If
CheckOC
strTemp(0) = Trim(str(Val(strTemp(1)) + Val(strTemp(2))))
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "總件數：" & "  " & strTemp(0)
Printer.CurrentX = 4500
Printer.CurrentY = iPrint
Printer.Print "無代理人件數：" & "  " & strTemp(1)
Printer.CurrentX = 9000
Printer.CurrentY = iPrint
If Trim(txt1(3)) = "1" Then
   Printer.Print "收文件數：" & "  " & strTemp(2)
Else
   Printer.Print "發文件數：" & "  " & strTemp(2)
End If
Printer.EndDoc
ShowPrintOk
End Sub

Private Sub PrintTitle()
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "代理人/新申請案排行榜") / 2)
Printer.CurrentY = iPrint
'Printer.Print Trim(Me.Tag) & " 代理人/新申請案排行榜"
Printer.Print GetTitleNick & "代理人/新申請案排行榜"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
If txt1(3) = "1" Then
    Printer.CurrentX = 7500 - (Printer.TextWidth("收文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "收文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
Else
    Printer.CurrentX = 7500 - (Printer.TextWidth("發文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))) / 2)
    Printer.CurrentY = iPrint
    Printer.Print "發文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
End If
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & txt1(1) & "－" & txt1(2)
'Add By Cheng 2002/02/19
Printer.CurrentX = 3500
Printer.CurrentY = iPrint
Printer.Print "國籍：" & txt1(8) & "－" & txt1(9)

Printer.CurrentX = 6800
Printer.CurrentY = iPrint
Select Case intPCaseKind
Case 1
     Printer.Print "系統類別：" & Me.txt1(0).Text
Case 2
     Printer.Print "系統類別：" & Me.txt1(0).Text
Case Else
End Select
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "國籍"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "排名"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "代理人名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "代理人編號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "CF件數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "FC件數"
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
End Sub

Private Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 2000
PLeft(2) = 3500
PLeft(3) = 9000
PLeft(4) = 11700
PLeft(5) = 13700
End Sub

Private Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
Printer.CurrentX = PLeft(1) + 600 - Printer.TextWidth(strTemp(1))
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print strTemp(3)
Printer.CurrentX = PLeft(4) + 600 - Printer.TextWidth(strTemp(4))
Printer.CurrentY = iPrint
Printer.Print strTemp(4)
Printer.CurrentX = PLeft(5) + 600 - Printer.TextWidth(strTemp(5))
Printer.CurrentY = iPrint
Printer.Print strTemp(5)
iPrint = iPrint + 300
End Sub

'申請人
Private Sub Process1()
Screen.MousePointer = vbHourglass
cnnConnection.Execute "DELETE FROM R050405_2 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
Select Case intPCaseKind
Case 1
      strSQL1 = strSQL1 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 1) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/10/19
      'Modify By Cheng 2002/02/19
'      strSQL1 = strSQL1 + " AND ((PA09>='" & txt1(1) & "' AND PA09<='" & txt1(2) & "') OR (CU10>='" & txt1(1) & "' AND CU10<='" & txt1(2) & "z')) "
        'Modify By Cheng 2003/11/11
'      strSQL1 = strSQL1 + " AND (PA09>='" & txt1(1) & "' AND PA09<='" & txt1(2) & "') And ((CU10>='" & txt1(8) & "' AND CU10<='" & txt1(9) & "z')) "
        If Trim(Me.txt1(1).Text) <> "" Then
            strSQL1 = strSQL1 & " And PA09>='" & Me.txt1(1).Text & "' "
        Else
            strSQL1 = strSQL1 & " And PA09=PA09 "
        End If
        If Trim(Me.txt1(2).Text) <> "" Then
            strSQL1 = strSQL1 & " And PA09<='" & Me.txt1(2).Text & "' "
        End If
        If Trim(Me.txt1(1).Text) <> "" Or Trim(Me.txt1(2).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
        End If
        If Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL1 = strSQL1 & " AND (CU10>='" & txt1(8) & "' AND CU10<='" & txt1(9) & "z') "
        ElseIf Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) = "" Then
            strSQL1 = strSQL1 & " AND CU10>='" & txt1(8) & "' ) "
        ElseIf Trim(Me.txt1(8).Text) = "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL1 = strSQL1 & " AND CU10<='" & txt1(9) & "z' "
        Else
            strSQL1 = strSQL1 & " AND CU10=CU10 "
        End If
        If Trim(Me.txt1(8).Text) <> "" Or Trim(Me.txt1(9).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label10 & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/10/19
        End If
Case 2
      strSQL2 = strSQL2 + " AND CP01 IN (" & SQLGrpStr(txt1(0), 2) & ") "
      pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/10/19
      'Modify By Cheng 2002/02/19
'      strSQL2 = strSQL2 + " AND ((TM10>='" & txt1(1) & "' AND TM10<='" & txt1(2) & "') OR (CU10>='" & txt1(1) & "' AND CU10<='" & txt1(2) & "z')) "
        'Modify By Cheng 2003/11/11
'      strSQL2 = strSQL2 + " AND (TM10>='" & txt1(1) & "' AND TM10<='" & txt1(2) & "') And ((CU10>='" & txt1(8) & "' AND CU10<='" & txt1(9) & "z')) "
        If Trim(Me.txt1(1).Text) <> "" Then
            strSQL2 = strSQL2 & " And TM10>='" & Me.txt1(1).Text & "' "
        Else
            strSQL2 = strSQL2 & " And TM10=TM10 "
        End If
        If Trim(Me.txt1(2).Text) <> "" Then
            strSQL2 = strSQL2 & " And TM10<='" & Me.txt1(2).Text & "' "
        End If
        If Trim(Me.txt1(1).Text) <> "" Or Trim(Me.txt1(2).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/10/19
        End If
        If Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL2 = strSQL2 & " AND (CU10>='" & txt1(8) & "' AND CU10<='" & txt1(9) & "z') "
        ElseIf Trim(Me.txt1(8).Text) <> "" And Trim(Me.txt1(9).Text) = "" Then
            strSQL2 = strSQL2 & " AND CU10>='" & txt1(8) & "' ) "
        ElseIf Trim(Me.txt1(8).Text) = "" And Trim(Me.txt1(9).Text) <> "" Then
            strSQL2 = strSQL2 & " AND CU10<='" & txt1(9) & "z' "
        Else
            strSQL2 = strSQL2 & " AND CU10=CU10 "
        End If
        If Trim(Me.txt1(8).Text) <> "" Or Trim(Me.txt1(9).Text) <> "" Then
            pub_QL05 = pub_QL05 & ";" & Label10 & txt1(8) & "-" & txt1(9) 'Add By Sindy 2010/10/19
        End If
Case Else
End Select
If Trim(txt1(3)) = "1" Then
   If Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(4))) & " "
      strSQL2 = strSQL2 + " AND CP05>=" & Val(ChangeTStringToWString(txt1(4))) & " "
   End If
   If Len(Trim(txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(5))) & " "
      strSQL2 = strSQL2 & " AND CP05<=" & Val(ChangeTStringToWString(txt1(5))) & " "
   End If
   If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";收文" & Label5 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/19
   End If
Else
   If Len(Trim(txt1(4))) <> 0 Then
      strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(4))) & " "
      strSQL2 = strSQL2 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(4))) & " "
   End If
   If Len(Trim(txt1(5))) <> 0 Then
      strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(5))) & " "
      strSQL2 = strSQL2 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(5))) & " "
   End If
   If Len(Trim(txt1(4))) <> 0 Or Len(Trim(txt1(5))) <> 0 Then
      pub_QL05 = pub_QL05 & ";發文" & Label5 & txt1(4) & "-" & txt1(5) 'Add By Sindy 2010/10/19
   End If
End If
'Modify By Cheng 2003/10/03
'Begin
'strSQL1 = strSQL1 + " AND (CP10='101' OR CP10='102' OR CP10='103' OR CP10='104' OR CP10='105') "
'2008/3/12 MODIFY BY SONIA 加CP09<'B'否則EPC指定國家子案也會計入
'strSQL1 = strSQL1 + " AND (CP10='101' OR CP10='102' OR CP10='103' OR CP10='104' OR CP10='105' Or CP10='109' Or CP10='110' Or CP10='112' Or CP10='113' Or CP10='114' Or CP10='115') "
strSQL1 = strSQL1 + " AND CP09<'B' AND (CP10='101' OR CP10='102' OR CP10='103' OR CP10='104' OR CP10='105' OR CP10='125' Or CP10='109' Or CP10='110' Or CP10='112' Or CP10='113' Or CP10='114' Or CP10='115') "
'End
'2008/3/12 MODIFY BY SONIA 加CP09<'B'否則EPC指定國家子案也會計入
'strSQL2 = strSQL2 + " AND (CP10='101') "
strSQL2 = strSQL2 + " AND CP09<'B' AND (CP10='101') "
CheckOC
Select Case intPCaseKind
Case 1
     strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,CU10,PA26,pa09,CP44,PA75 FROM PATENT,CASEPROGRESS,CUSTOMER WHERE " & SQLNewFag("PA26", "CU") & " AND CP01=pa01(+) AND CP02=pa02(+) AND CP03=pa03(+) AND CP04=pa04(+) AND CP57 IS NULL" & strSQL1
Case 2
     strSql = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,CU10,TM23,tm10,CP44,TM44 FROM TRADEMARK,CASEPROGRESS,CUSTOMER WHERE " & SQLNewFag("TM23", "CU") & " AND CP01=tm01(+) AND CP02=tm02(+) AND CP03=tm03(+) AND CP04=tm04(+) AND CP57 IS NULL " & strSQL2
'strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,NA03,CP44,PA75,NA01,PA09 FROM NATION,PATENT,CASEPROGRESS WHERE PA09=na01(+) AND cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND CP57 IS NULL " & StrSQL1
'strSQL = strSQL + " union all select CP01||'-'||CP02||'-'||CP03||'-'||CP04,NA03,CP44,TM44,NA01,TM10 FROM NATION,TRADEMARK,CASEPROGRESS WHERE TM10=na01(+) AND cp01=TM01(+) AND cp02=TM02(+) AND cp03=TM03(+) AND cp04=TM04(+) AND CP57 IS NULL " & StrSQL2
Case Else
End Select
strSql = strSql + " ORDER BY 2 "
'Select Case intPCaseKind
'Case 1
'      strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,F2.FA10,CP44,PA75,F1.FA10,PA09 FROM FAGENT F1,FAGENT F2,PATENT,CASEPROGRESS " & _
'         " WHERE CP01=pa01(+) AND CP02=pa02(+) AND CP03=pa03(+) AND CP04=pa04(+) AND CP57 IS NULL AND " & SQLNewFag("PA75", "F1.FA") & " AND " & SQLNewFag("CP44", "F2.FA") & strSQL1
'Case 2
'      strSQL = "SELECT CP01||'-'||CP02||'-'||CP03||'-'||CP04,F2.FA10,CP44,TM44,F1.FA10,TM10 FROM FAGENT F1,FAGENT F2,TRADEMARK,CASEPROGRESS " & _
'         " WHERE CP01=tm01(+) AND CP02=tm02(+) AND cP03=tm03(+) AND CP04=tm04(+) AND CP57 IS NULL AND " & SQLNewFag("TM44", "F1.FA") & " AND " & SQLNewFag("CP44", "F2.FA") & strSQL2
'Case Else
'End Select


k = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        For i = 0 To 5
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        strTemp(0) = IIf(Trim(GetPrjSalesNM(Get050405(GetNewFagent(strTemp(2))))) = "", Get050405(GetNewFagent(strTemp(2))), GetPrjSalesNM(Get050405(GetNewFagent(strTemp(2)))))
        'Modify By Cheng 2003/10/06
        '若代理人國籍為台灣或大陸, 則代理人名稱中-->英-->日, 其他則英-->中-->日
        'Begin
        If strTemp(1) < "010" Or strTemp(1) = "020" Then
            strTemp(1) = IIf(CheckStr(GetPrjPeople1(GetNewFagent(strTemp(2)), "1")) = "", strTemp(2), GetPrjPeople1(GetNewFagent(strTemp(2)), "1"))
        Else
            strTemp(1) = IIf(CheckStr(GetPrjPeople1(GetNewFagent(strTemp(2)), "2")) = "", strTemp(2), GetPrjPeople1(GetNewFagent(strTemp(2)), "2"))
        End If
        strTemp(2) = GetNewFagent(strTemp(2))
        strTemp(3) = "1"
        strTemp(6) = ""
        strTemp(7) = IIf(Trim(GetPrjNationName(GetPrjNationNumber1(GetNewFagent(strTemp(2))))) = "", GetPrjNationNumber1(GetNewFagent(strTemp(2))), GetPrjNationName(GetPrjNationNumber1(GetNewFagent(strTemp(2)))))
        If Len(Trim(strTemp(4))) = 0 And Len(Trim(strTemp(5))) = 0 Then
           strTemp(6) = "*"
        End If
        strSql = "INSERT INTO R050405_2 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & ",'" & ChgSQL(strTemp(6)) & "','" & ChgSQL(strTemp(7)) & "','" & strUserNum & "') "
        cnnConnection.Execute strSql
        DoEvents
        adoRecordset.MoveNext
    Loop
Else
    InsertQueryLog (0) 'Add By Sindy 2010/10/19
    ShowNoData
    Screen.MousePointer = vbDefault
    Exit Sub
End If
CheckOC
strSql = "SELECT R020001,R020002,R020003,SUM(R020004),R020005,R020006,'" & strUserNum & "' FROM R050405_2 WHERE ID='" & strUserNum & "'GROUP BY R020001,R020002,R020003,R020005,R020006,'" & strUserNum & "' "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/10/19
    adoRecordset.MoveFirst
    cnnConnection.Execute "DELETE FROM R050405_2 WHERE ID='" & strUserNum & "'"
    Do While adoRecordset.EOF = False
        For i = 0 To 5
            strTemp(i) = CheckStr(adoRecordset.Fields(i))
        Next i
        strSql = "INSERT INTO R050405_2 VALUES('" & ChgSQL(strTemp(0)) & "','" & ChgSQL(strTemp(1)) & "','" & ChgSQL(strTemp(2)) & "'," & Val(strTemp(3)) & ",'" & ChgSQL(strTemp(4)) & "','" & ChgSQL(strTemp(5)) & "','" & strUserNum & "') "
        cnnConnection.Execute strSql
        adoRecordset.MoveNext
    Loop
Else
    InsertQueryLog (0)  'Add By Sindy 2010/10/19
End If
CheckOC
PrintData1

Screen.MousePointer = vbDefault
End Sub

Private Sub PrintData1()
'strSQL = "SELECT * FROM R050405_2 WHERE ID='" & strUserNum & "' ORDER BY R020004 DESC "
'Modify By Cheng 2002/02/19
'strSQL = "SELECT R020001,R020002,R020003,SUM(R020004),R020006 FROM R050405_2 WHERE ID='" & strUserNum & "' AND R020003 IS NOT NULL GROUP BY R020001,R020002,R020003,R020006 ORDER BY SUM(R020004) DESC "
strSql = "SELECT R020001,R020002,R020003,SUM(R020004),R020006 FROM R050405_2 WHERE ID='" & strUserNum & "' AND R020003 IS NOT NULL GROUP BY R020001,R020002,R020003,R020006 ORDER BY SUM(R020004) DESC "
CheckOC
Page = 1
StrTest99 = "      "
iY = 0
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        PrintTitle1
        Do While .EOF = False
            For i = 1 To 5
                strTemp(i) = CheckStr(.Fields(i - 1))
            Next i
            strTemp(0) = strTemp(1)
            If StrTest99 <> strTemp(4) Then
                iY = iY + 1
                strTemp(1) = Trim(str(iY))
                StrTest99 = strTemp(4)
            Else
                strTemp(1) = Trim(str(iY))
            End If
            strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 12), vbUnicode)
            strTemp(2) = StrConv(MidB(StrConv(strTemp(2), vbFromUnicode), 1, 36), vbUnicode)
            If Len(txt1(7)) <> 0 Then
               If iY > Val(txt1(7)) Then
                  GoTo OutPrintDoc1
               End If
            End If
            If iPrint > 10000 Then
                Page = Page + 1
                Printer.CurrentX = 500
                Printer.CurrentY = iPrint
                Printer.Print String(250, "-")
                Printer.NewPage
                PrintTitle1
            End If
            PrintDatil1
            .MoveNext
        Loop
    End With
End If
OutPrintDoc1:
CheckOC
strTemp(2) = "0"
strTemp(1) = "0"
strTemp(0) = "0"
strSql = "SELECT SUM(R020004) FRoM R050405_2 WHERE R020005='*' AND ID='" & strUserNum & "'"
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    strTemp(1) = CheckStr(adoRecordset.Fields(0))
End If
CheckOC
strSql = "SELECT SUM(R020004) FRoM R050405_2 WHERE ID='" & strUserNum & "' AND R020003 IS NOT NULL "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    strTemp(2) = CheckStr(adoRecordset.Fields(0))
End If
CheckOC
strTemp(0) = strTemp(2)
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "總件數：" & "  " & strTemp(0)
Printer.CurrentX = 4500
Printer.CurrentY = iPrint
Printer.Print "無代理人件數：" & "  " & strTemp(1)
Printer.CurrentX = 9000
Printer.CurrentY = iPrint
If Trim(txt1(3)) = "1" Then
   Printer.Print "收文件數：" & "  " & strTemp(2)
Else
   Printer.Print "發文件數：" & "  " & strTemp(2)
End If
Printer.EndDoc
ShowPrintOk

End Sub

Private Sub PrintTitle1()
GetPleft1
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 7500 - (Printer.TextWidth(GetTitleNick & "申請人/新申請案排行榜") / 2)
Printer.CurrentY = iPrint
Printer.Print GetTitleNick & "申請人/新申請案排行榜"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
If txt1(3) = "1" Then
Printer.CurrentX = 7500 - (Printer.TextWidth("收文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))) / 2)
Printer.CurrentY = iPrint
    Printer.Print "收文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
Else
Printer.CurrentX = 7500 - (Printer.TextWidth("發文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))) / 2)
Printer.CurrentY = iPrint
    Printer.Print "發文日期：" & Format(ChangeTStringToTDateString(txt1(4)), "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(5))
End If
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(strSrvDate(2), "##/##/##")
iPrint = iPrint + 300
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print "區別：" & StrTemp5(0) & "－" & StrTemp5(1)
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "申請國家：" & txt1(1) & "－" & txt1(2)
'Add By Cheng 2002/02/19
Printer.CurrentX = 3500
Printer.CurrentY = iPrint
Printer.Print "國籍：" & txt1(8) & "－" & txt1(9)

Printer.CurrentX = 6800
Printer.CurrentY = iPrint
Printer.Print "系統類別:" & ChgNewStr(txt1(0))
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
Printer.Font.Underline = True
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "智權人員"
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "排名"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "申請人名稱"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "申請人編號"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "申請人國籍"
Printer.Font.Underline = False
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(250, "-")
iPrint = iPrint + 300
End Sub

Private Sub GetPleft1()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 2000
PLeft(2) = 3500
PLeft(3) = 10000
PLeft(4) = 12500
PLeft(5) = 15000
End Sub

Private Sub PrintDatil1()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(0)
Printer.CurrentX = PLeft(1) + 600 - Printer.TextWidth(strTemp(1))
Printer.CurrentY = iPrint
Printer.Print strTemp(1)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(2)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print strTemp(3)
Printer.CurrentX = PLeft(4) + 600 - Printer.TextWidth(strTemp(4))
Printer.CurrentY = iPrint
Printer.Print strTemp(4)
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print strTemp(5)
iPrint = iPrint + 300
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
Select Case intPCaseKind
Case 1
     txt1(0) = "P,CFP,FCP"
Case 2
     txt1(0) = "T,CFT,FCT,TF"
Case Else
End Select
'Me.Tag = StrStartSystemByNick
'StrTest = StrTest2
'strTemp1 = Split(Replace(UCase(StrTest), ",,", ""), ",")
'strTemp2 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
'For i = 0 To UBound(strTemp1)
'    S = 0
'    For j = 0 To UBound(strTemp2)
'        If strTemp2(j) = strTemp1(i) Then
'            S = 1
'            Exit For
'        End If
'    Next j
'    If S = 0 Then
'        StrTest = Replace(StrTest, "," & strTemp1(i) & ",", "")
'    End If
'Next i
'txt1(0) = StrTest
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050405 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub


Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '系統類別
     'strTemp1 = Split(Replace(UCase(GetSystemKindByNick), ",,", ""), ",")
     'strTemp2 = Split(Replace(UCase(txt1(0)), ",,", ""), ",")
     'For i = 0 To UBound(strTemp2)
        'S = 0
        'For j = 0 To UBound(strTemp1)
            'If strTemp1(j) = strTemp2(i) Then
                'S = 1
                'Exit For
            'End If
        'Next j
        'If S = 0 Then
            'S = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            'txt1(0).SetFocus
            'txt1(0).SelStart = 0
            'txt1(0).SelLength = Len(txt1(0))
            'Exit Sub
        'End If
     'Next i
Case 2 '申請國家
     If RunNick(txt1(1), txt1(2)) Then
         txt1(1).SetFocus
         txt1_GotFocus (1)
         Exit Sub
      End If
Case 3 '列印別
     Select Case txt1(3)
     Case "1", "2", "", " "
     Case Else
          s = MsgBox("列印別只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(3).SetFocus
          txt1(3).SelStart = 0
          txt1(3).SelLength = Len(txt1(3))
          Exit Sub
     End Select
Case 4, 5 '日期
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
      Exit Sub
   End If
   If Index = 5 Then
     If RunNick(txt1(4), txt1(5)) Then
         txt1(4).SetFocus
         txt1_GotFocus (4)
         Exit Sub
       End If
    End If
Case 6 '列印對象
     Select Case txt1(6)
     Case "1", "2", "", " "
     Case Else
'          s = MsgBox("排行條件只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          s = MsgBox("列印對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(6).SetFocus
          txt1(6).SelStart = 0
          txt1(6).SelLength = Len(txt1(6))
          Exit Sub
     End Select
Case 7 '列印名次
     If Trim(txt1(7)) <> "" Then
         If IsNumeric(txt1(7)) = False Then
                s = MsgBox("列印名次只可輸入數字!!", , "USER 輸入錯誤")
                txt1(7).SetFocus
                txt1_GotFocus (7)
                Exit Sub
         End If
     End If
Case 9 '國籍
     If RunNick(txt1(8), txt1(9)) Then
         txt1(8).SetFocus
         txt1_GotFocus (8)
         Exit Sub
      End If
Case 10 '排行條件
      If Me.txt1(6).Text = "1" Then
        Select Case txt1(10)
        Case "1", "2", "", " "
        Case Else
   '          s = MsgBox("排行條件只能輸入 1 或 2 !!", , "USER 輸入錯誤")
             s = MsgBox("列印對象只能輸入 1 或 2 !!", , "USER 輸入錯誤")
             txt1(10).SetFocus
             txt1(10).SelStart = 0
             txt1(10).SelLength = Len(txt1(10))
             Exit Sub
        End Select
      End If
Case Else
End Select
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
KeyAscii = UpperCase(KeyAscii)
End Sub


Private Function Get050405(ByRef Strindex As String) As String
If Len(Trim(Strindex)) = 0 Then
    Get050405 = ""
    Exit Function
End If
strSql = "SELECT CU13 FROM CUSTOMER WHERE CU01='" & Mid(Strindex, 1, 8) & "' AND CU02='" & Mid(Strindex, 9, 1) & "' "
CheckOC3
AdoRecordSet3.CursorLocation = adUseClient
AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
    Get050405 = CheckStr(AdoRecordSet3.Fields(0))
Else
    Get050405 = ""
End If
CheckOC3
End Function


