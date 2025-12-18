VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm060305 
   BorderStyle     =   1  '單線固定
   Caption         =   "年費通知月報表"
   ClientHeight    =   2964
   ClientLeft      =   3192
   ClientTop       =   1968
   ClientWidth     =   6276
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2964
   ScaleWidth      =   6276
   Begin VB.Frame Frame1 
      Caption         =   "設定地址條"
      Height          =   660
      Left            =   90
      TabIndex        =   13
      Top             =   2190
      Width           =   3645
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   765
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   240
         Width           =   2670
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   2
         Left            =   105
         TabIndex        =   14
         Top             =   255
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   5190
      TabIndex        =   5
      Top             =   36
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   4410
      TabIndex        =   4
      Top             =   36
      Width           =   756
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2388
      MaxLength       =   7
      TabIndex        =   2
      Top             =   912
      Width           =   1230
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   1032
      MaxLength       =   7
      TabIndex        =   1
      Top             =   912
      Width           =   1230
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   1584
      MaxLength       =   9
      TabIndex        =   0
      Top             =   552
      Width           =   1230
   End
   Begin MSForms.Label lbl1 
      Height          =   225
      Left            =   2880
      TabIndex        =   15
      Top             =   600
      Width           =   3315
      VariousPropertyBits=   27
      Size            =   "5847;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line1 
      X1              =   1656
      X2              =   3231
      Y1              =   1056
      Y2              =   1056
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Novartis International AG."
      Height          =   180
      Index           =   8
      Left            =   810
      TabIndex        =   12
      Top             =   1860
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "CIBA SPECIALTY CHEMICALS HOLDING INC PATENT DEPT."
      Height          =   180
      Index           =   7
      Left            =   810
      TabIndex        =   11
      Top             =   1515
      Width           =   4770
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y49575"
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   10
      Top             =   1860
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Y45697"
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   9
      Top             =   1515
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "目前適用對象："
      Height          =   180
      Index           =   2
      Left            =   84
      TabIndex        =   8
      Top             =   1260
      Width           =   1392
   End
   Begin VB.Label Label1 
      Caption         =   "年費期限："
      Height          =   180
      Index           =   1
      Left            =   84
      TabIndex        =   7
      Top             =   972
      Width           =   936
   End
   Begin VB.Label Label1 
      Caption         =   "代理人(申請人)："
      Height          =   180
      Index           =   0
      Left            =   84
      TabIndex        =   6
      Top             =   612
      Width           =   1392
   End
End
Attribute VB_Name = "frm060305"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/15 Form2.0已修改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'Memo by Morgan2010/12/27 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/13 日期欄已修改
Option Explicit

Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, strTemp3(0 To 1) As String
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 26) As String, StrTemp99(0 To 15) As String
Dim PLeft(0 To 25) As Integer, strTemp1 As Variant, strTemp2 As Variant, Bol1 As Boolean
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
'Add By Cheng 2002/12/30
Const m_PrintLeftPos = 500 '列印X軸起點
'Add By Cheng 2003/01/28
Dim m_OriPrinterName As String, SeekPrint As Integer, SeekPrintL As Integer
Dim m_AddrList As String
'Add by Morgan 2011/3/15
Dim strPrinter As String
Dim m_LetterLanguage As String 'Add By Sindy 2015/9/21


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0 '確定
    'Add By Cheng 2002/09/16
    blnClkSure = False
     Printer.Orientation = 2
     DoEvents
     If Len(txt1(0)) = 0 Then
         s = MsgBox("代理人(申請人)不可空白!!", , "USER 輸入錯誤")
         txt1(0).SetFocus
         Exit Sub
     Else
         'Add By Cheng 2002/09/16
         If Mid(UCase(txt1(0)), 1, 1) = "X" Then
            strSql = "select nvl(cu05||cu88||cu89||cu90,nvl(cu04,cu06)) from customer where cu01='" & Mid(GetNewFagent(txt1(0)), 1, 8) & "' and cu02='" & Mid(GetNewFagent(txt1(0)), 9, 1) & "' "
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                lbl1.Caption = CheckStr(adoRecordset.Fields(0))
            Else
                MsgBox "代理人(申請人)代號錯誤，請重新輸入 !", vbCritical
                lbl1.Caption = ""
                txt1(0).SetFocus
                txt1_GotFocus 0
                CheckOC
                Exit Sub
            End If
            CheckOC
         ElseIf Mid(UCase(txt1(0)), 1, 1) = "Y" Then
            strSql = "select nvl(fa05||fa63||fa64||fa65,nvl(fa04,fa06)) from fagent where fa01='" & Mid(GetNewFagent(txt1(0)), 1, 8) & "' and fa02='" & Mid(GetNewFagent(txt1(0)), 9, 1) & "' "
            CheckOC
            adoRecordset.CursorLocation = adUseClient
            adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
                lbl1.Caption = CheckStr(adoRecordset.Fields(0))
            Else
                MsgBox "代理人(申請人)代號錯誤，請重新輸入 !", vbCritical
                lbl1.Caption = ""
                Me.txt1(0).SetFocus
                txt1_GotFocus 0
                CheckOC
                Exit Sub
            End If
            CheckOC
         Else
            MsgBox "代理人(申請人)代號錯誤，請重新輸入 !", vbCritical
            lbl1.Caption = ""
            txt1(0).SetFocus
            txt1_GotFocus 0
            Exit Sub
         End If
         
         If Len(txt1(2)) = 0 Then
             s = MsgBox("年費期限區間不可空白!!", , "USER 輸入錯誤")
             
             If Len(txt1(1)) = 0 Then txt1(1).SetFocus
             Exit Sub
         Else
            'Add By Cheng 2002/03/20
            If PUB_CheckKeyInDate(Me.txt1(1)) = -1 Then
               Me.txt1(1).SetFocus
               txt1_GotFocus 1
               Exit Sub
            End If
            If PUB_CheckKeyInDate(Me.txt1(2)) = -1 Then
               Me.txt1(2).SetFocus
               txt1_GotFocus 2
               Exit Sub
            End If
            'Add By Cheng 2002/09/16
            If Me.txt1(1).Text <> "" And Me.txt1(2).Text <> "" Then
               If Val(Me.txt1(1).Text) > Val(Me.txt1(2).Text) Then
                  MsgBox "年費期限範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
             
             Screen.MousePointer = vbHourglass
             Me.Enabled = False
             ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/7 清除查詢印表記錄檔欄位
             Process
             Me.Enabled = True
             Screen.MousePointer = vbDefault
         End If
     End If
Case 1 '結束
   Me.Enabled = False
   
   'Move to Unload by Morgan 2004/10/26
'    'Add By Cheng 2003/01/29
'    '列印地址條
'    PUB_PrintAddressList strUserNum, Me.Combo1.Text
'    '刪除地址條列表資料
'    PUB_DeleteAddressList strUserNum
'    '初始化序號
'    pub_AddressListSN = 0
'    'Add By Cheng 2003/02/05
'    '若印表機變動, 則更新列印設定
'    If Me.Combo1.Text <> Me.Combo1.Tag Then
'        PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
'    End If
   '2004/10/26 end
   
    Unload Me
Case Else
End Select
End Sub

Sub Process()
Dim bolCiba As Boolean

pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(0) & lbl1 'Add By Sindy 2010/12/7
'頭
'Modify by Morgan 2011/5/26 +CU102,FA70
If Mid(UCase(txt1(0)), 1, 1) = "X" Then
    strSql = "SELECT CU05,CU88,CU89,CU90,CU65,CU66,CU67,CU68,CU69,CU24,CU25,CU26,CU27,CU28,NA04,CU59,CU62,CU102 FROM CUSTOMER,NATION WHERE CU01='" & Mid(GetNewFagent(txt1(0)), 1, 8) & "' AND CU02='" & Mid(GetNewFagent(txt1(0)), 9, 1) & "' AND CU10=NA01(+) "
Else
    If Mid(UCase(txt1(0)), 1, 1) = "Y" Then
        'Add by Morgan 2010/2/25 Ciba的收件人改抓 Y5296500
        If GetNewFagent(txt1(0)) = "Y45697000" Then
            bolCiba = True
            strSql = "SELECT FA05,FA63,FA64,FA65,FA32,FA33,FA34,FA35,FA36,FA18,FA19,FA20,FA21,FA22,NA04,'','',FA70 FROM FAGENT,NATION WHERE FA01='Y5296500' AND FA02='0' AND FA10=NA01(+) "
        Else
        'end 2010/2/25
            strSql = "SELECT FA05,FA63,FA64,FA65,FA32,FA33,FA34,FA35,FA36,FA18,FA19,FA20,FA21,FA22,NA04,'','',FA70 FROM FAGENT,NATION WHERE FA01='" & Mid(GetNewFagent(txt1(0)), 1, 8) & "' AND FA02='" & Mid(GetNewFagent(txt1(0)), 9, 1) & "' AND FA10=NA01(+) "
        End If
    End If
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    With adoRecordset
        .MoveFirst
        Do While .EOF = False
            For i = 0 To 16
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            strTemp(26) = CheckStr(.Fields(17)) 'Add by Morgan 2011/5/26
            .MoveNext
        Loop
    End With
End If
CheckOC
Page = 1
strSQL1 = ""
If Len(txt1(1)) <> 0 Then
    'Modify By Cheng 2002/12/20
    '判斷法定期限
    'strSQL1 = strSQL1 + " AND NP08>=" & Val(ChangeTStringToWString(txt1(1))) & ""
    strSQL1 = strSQL1 + " AND NP09>=" & Val(ChangeTStringToWString(txt1(1))) & ""
End If
If Len(txt1(2)) <> 0 Then
    'Modify By Cheng 2002/12/20
    '判斷法定期限
'    strSQL1 = strSQL1 + " AND NP08<=" & Val(ChangeTStringToWString(txt1(2))) & " "
    strSQL1 = strSQL1 + " AND NP09<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
If Len(txt1(1)) <> 0 Or Len(txt1(2)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(1) & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/12/7
End If
strSQL1 = strSQL1 + " AND PA01='FCP' "

'Add by Morgan 2004/8/10  加控制閉卷不印
strSQL1 = strSQL1 & " AND PA57 IS NULL "
'Add by Morgan 2006/5/8 加控制期限有管制的
strSQL1 = strSQL1 & " AND NP06 IS NULL "

If Mid(UCase(txt1(0)), 1, 1) = "X" Then
    strSQL1 = strSQL1 + " AND (PA26='" & GetNewFagent(txt1(0)) & "' OR PA27='" & GetNewFagent(txt1(0)) & "' OR PA28='" & GetNewFagent(txt1(0)) & "' OR PA29='" & GetNewFagent(txt1(0)) & "' OR PA30='" & GetNewFagent(txt1(0)) & "' ) "
    'Modify By Cheng 2002/12/20
    '顯示法定期限非本所期限
'    strSQL = "SELECT PA22,PA11,PA77,NP08,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA48,PA72,PA08,PA09 FROM PATENT,NEXTPROGRESS WHERE NP07=605 AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) " & strSQL1
   strSql = "SELECT PA22,PA11,PA77,NP09,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA48,PA72,PA08,PA09,PA01,PA02,PA03,PA04 FROM PATENT,NEXTPROGRESS WHERE NP07=605 AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) " & strSQL1
    'Add By Cheng 2002/12/20
    strSql = strSql & " ORDER BY 4,5 "
Else
    If Mid(UCase(txt1(0)), 1, 1) = "Y" Then
        'Modify by Morgan 2004/9/1 也抓年費代理人
        'strSQL1 = strSQL1 + " AND PA75='" & GetNewFagent(Txt1(0)) & "' "
        strSQL1 = strSQL1 + " AND ( PA75='" & GetNewFagent(txt1(0)) & "' or PA76='" & GetNewFagent(txt1(0)) & "' )"
        
        'Modify By Cheng 2002/12/20
        '顯示法定期限非本所期限
'        strSQL = "SELECT PA22,PA11,PA77,NP08,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA48,PA72,PA08,PA09 FROM PATENT,nextprogress WHERE NP07=605 AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) " & strSQL1
        strSql = "SELECT PA22,PA11,PA77,NP09,PA01||'-'||PA02||'-'||PA03||'-'||PA04,PA48,PA72,PA08,PA09,PA01,PA02,PA03,PA04 FROM PATENT,nextprogress WHERE NP07=605 AND PA01=NP02(+) AND PA02=NP03(+) AND PA03=NP04(+) AND PA04=NP05(+) " & strSQL1
        'Add By Cheng 2002/12/20
        strSql = strSql & " ORDER BY 4,5 "
    End If
End If
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/12/7
    'Add By Cheng 2003/01/08
    MsgBox "請確認列印紙張!!!", vbExclamation + vbOKOnly
    With adoRecordset
        .MoveFirst
        
        m_LetterLanguage = PUB_GetLanguage("" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value) 'Add By Sindy 2015/9/21
        
        If bolCiba Then
            m_AddrList = "Y52965000"
        End If
'        'Add By Sindy 2015/9/21 日文定稿才要印地址條
'        If m_LetterLanguage = "3" Or Val(外專開窗信函啟用日) >= Val(strSrvDate(1)) Then
'        '2015/9/21 END
           'Add By Cheng 2003/01/29
           '新增地址條列表資料
           pub_AddressListSN = pub_AddressListSN + 1
           'Modify By Cheng 2003/02/07
           '加傳入綠皮貼紙的份數
   '        PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN
           'Modify by Morgan 2010/2/25 加傳AL08
           'PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN, "0"
           If bolCiba Then
               m_AddrList = "Y52965000"
           Else
               PUB_AddNewAddressList strUserNum, "" & .Fields("PA01").Value, "" & .Fields("PA02").Value, "" & .Fields("PA03").Value, "" & .Fields("PA04").Value, "" & pub_AddressListSN, "0", "605"
           End If
'        End If
        
        PrintTitle
        Do While .EOF = False
            strSql = "SELECT NA21,NA23,NA25 FROM NATION WHERE NA01='" & CheckStr(.Fields(8)) & "' "
            CheckOC2
            adoRecordset1.CursorLocation = adUseClient
            adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
                Select Case Val(CheckStr(.Fields(7)))
                Case 1
                     strTemp1 = Split(Replace(CheckStr(.Fields(6)), ",,", ""), ",")
                     strTemp2 = Split(Replace(CheckStr(adoRecordset1.Fields(0)), ",,", ""), ",")
                Case 2
                     strTemp1 = Split(Replace(CheckStr(.Fields(6)), ",,", ""), ",")
                     strTemp2 = Split(Replace(CheckStr(adoRecordset1.Fields(1)), ",,", ""), ",")
                Case 3
                     strTemp1 = Split(Replace(CheckStr(.Fields(6)), ",,", ""), ",")
                     strTemp2 = Split(Replace(CheckStr(adoRecordset1.Fields(2)), ",,", ""), ",")
                Case Else
                End Select
                If Len(CheckStr(.Fields(6))) = 0 Then
                    i = 0
                Else
                    For i = UBound(strTemp1) To 0 Step -1
                        If Len(strTemp1(i)) <> 0 Then
                            Exit For
                        End If
                    Next i
                    i = i + 1
                End If
                If UBound(strTemp2) < 0 Then
                    strTemp3(1) = ""
                Else
                    If i = 0 Then
                        strTemp3(1) = strTemp2(0)
                    Else
                        For j = 0 To UBound(strTemp2)
                            If strTemp1(i - 1) = strTemp2(j) Then
                                strTemp3(1) = strTemp2(j + 1)
                                Exit For
                            End If
                        Next j
                    End If
                End If
            End If
            strTemp3(0) = Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY")
            'Add By Cheng 2002/12/30
            '若有繳年費年度
            If strTemp3(1) <> "" Then strTemp3(1) = ChgYearFormat(strTemp3(1))
            For i = 0 To 8
                strTemp(17 + i) = CheckStr(.Fields(i))
            Next i
            strTemp(17) = StrToStr(strTemp(17), 4)
            strTemp(18) = StrToStr(strTemp(18), 5)
            'Modify By Cheng 2002/12/20
            '只顯示日的部分
'            strTemp(20) = ChangeWStringToWDateString(strTemp(20))
            strTemp(20) = Right(strTemp(20), 2)
            'Add By Cheng 2002/12/30
            '本所案號若後面序號為"-0-00"者不印出來
            strTemp(21) = Replace(strTemp(21), "-0-00", "")
            PrintDatil
            If iPrint > 13500 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle1
            End If
            .MoveNext
        Loop
    End With
Else
    InsertQueryLog (0) 'Add By Sindy 2010/12/7
    ShowNoData
    Exit Sub
End If
PrintEnd
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd()
If iPrint >= 11000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1
End If
iPrint = iPrint + 600
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 2500
Printer.CurrentX = 2500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "Service fee"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 4700
Printer.CurrentX = 4700 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "Official fee"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 7500
Printer.CurrentX = 7500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "Total"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 9500
Printer.CurrentX = 9500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "Total"
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 3000
Printer.CurrentX = 3000 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "(NT$)"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 5200
Printer.CurrentX = 5200 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "(NT$)"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 7500
Printer.CurrentX = 7500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "(NT$)"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 9500
Printer.CurrentX = 9500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "(USD)"
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "Annuity per annum："
iPrint = iPrint + 300
'Modify By Cheng 2002/12/20
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("1st to 5th")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("1st to 3th")
Printer.CurrentX = 300 + 1814 + m_PrintLeftPos - Printer.TextWidth("1st to 3th")
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/20
'Printer.Print "1st to 5th"
Printer.Print "1st to 3th"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 3000
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 3000 + m_PrintLeftPos
Printer.CurrentX = 3000 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(0), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(0), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 5200
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 5200 + m_PrintLeftPos
Printer.CurrentX = 5200 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(1), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(1), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 7500
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 7500 + m_PrintLeftPos
Printer.CurrentX = 7500 + m_PrintLeftPos + 600 - Printer.TextWidth(Format(StrTemp99(1), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(2), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 9500
Printer.CurrentX = 9500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(3), "###,###,###.00")
iPrint = iPrint + 300
'Modify By Cheng 2002/12/20
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("6th to 10th")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("4th to 6th")
Printer.CurrentX = 300 + 1814 + m_PrintLeftPos - Printer.TextWidth("4th to 6th")
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/20
'Printer.Print "6th to 10th"
Printer.Print "4th to 6th"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 3000
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 3000 + m_PrintLeftPos
Printer.CurrentX = 3000 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(4), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(4), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 5200
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 5200 + m_PrintLeftPos
Printer.CurrentX = 5200 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(5), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(5), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 7500
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 7500 + m_PrintLeftPos
Printer.CurrentX = 7500 + m_PrintLeftPos + 600 - Printer.TextWidth(Format(StrTemp99(6), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(6), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 9500
Printer.CurrentX = 9500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(7), "###,###,###.00")
iPrint = iPrint + 300
'Modify By Cheng 2002/12/20
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("11th to 15th")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("7th to 9th")
Printer.CurrentX = 300 + 1814 + m_PrintLeftPos - Printer.TextWidth("7th to 9th")
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/20
'Printer.Print "11th to 15th"
Printer.Print "7th to 9th"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 3000
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 3000 + m_PrintLeftPos
Printer.CurrentX = 3000 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(8), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(8), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 5200
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 5200 + m_PrintLeftPos
Printer.CurrentX = 5200 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(9), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(9), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 7500
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 7500 + m_PrintLeftPos
Printer.CurrentX = 7500 + m_PrintLeftPos + 600 - Printer.TextWidth(Format(StrTemp99(10), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(10), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 9500
Printer.CurrentX = 9500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(11), "###,###,###.00")
iPrint = iPrint + 300
'Modify By Cheng 2002/12/20
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("16th to 20th")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300 + 1814 - Printer.TextWidth("10th to 20th")
Printer.CurrentX = 300 + 1814 + m_PrintLeftPos - Printer.TextWidth("10th to 20th")
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/20
'Printer.Print "16th to 20th"
Printer.Print "10th to 20th"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 3000
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 3000 + m_PrintLeftPos
Printer.CurrentX = 3000 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(12), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(12), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 5200
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 5200 + m_PrintLeftPos
Printer.CurrentX = 5200 + m_PrintLeftPos + 800 - Printer.TextWidth(Format(StrTemp99(13), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(13), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 7500
'Modify by Morgan 2004/8/10   靠右對齊
'Printer.CurrentX = 7500 + m_PrintLeftPos
Printer.CurrentX = 7500 + m_PrintLeftPos + 600 - Printer.TextWidth(Format(StrTemp99(14), "###,###,###"))
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(14), "###,###,###")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 9500
Printer.CurrentX = 9500 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print Format(StrTemp99(15), "###,###,###.00")
iPrint = iPrint + 600

'Add by Morgan 2005/4/11
If iPrint >= 13200 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle1 True
End If
'2005/4/11

'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Font.Size = 12
Printer.Print "Could you please let us have instructions for paying the above patents in due course. "
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/24
'Printer.Print "We await your instructions "
Printer.Print "We await your instructions with kind regards. "
'iPrint = iPrint + 600
'Printer.CurrentX = 300
'Printer.CurrentY = iPrint
'Printer.Print "With kind regards. "

iPrint = iPrint + 600
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 8000
'Modify by Morgan 2004/1/14
'Printer.CurrentX = 8000 + m_PrintLeftPos
Printer.CurrentX = 8000 + m_PrintLeftPos - 3000

Printer.CurrentY = iPrint
Printer.Print "Best regards,"
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 8000
'Modify by Morgan 2004/1/14
'Printer.CurrentX = 8000 + m_PrintLeftPos
Printer.CurrentX = 8000 + m_PrintLeftPos - 3000

Printer.CurrentY = iPrint
Printer.Print "Tai E International"
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 8000
'Modify by Morgan 2004/1/14
'Printer.CurrentX = 8000 + m_PrintLeftPos
Printer.CurrentX = 8000 + m_PrintLeftPos - 3000

Printer.CurrentY = iPrint
Printer.Print "Patent & Law Office"
iPrint = iPrint + 600
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Font.Size = 11
Printer.Print "CYT/dy"
End Sub

Sub PrintDatil()
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print strTemp(17)
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print strTemp(18)
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print strTemp(19)
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print strTemp(20)
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print strTemp3(1)
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print strTemp(21)
iPrint = iPrint + 300
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print strTemp(22)
iPrint = iPrint + 600
End Sub

Sub GetPleft()
Erase PLeft
'Modify By Cheng 2002/12/30
'PLeft(0) = 300
'PLeft(1) = 1300
'PLeft(2) = 3000
'PLeft(3) = 7500 - 700
'PLeft(4) = 8700 - 500
'PLeft(5) = 9500 - 500
'PLeft(6) = 3000
PLeft(0) = 300 + 500
PLeft(1) = 1300 + 500
PLeft(2) = 3000 + 500
PLeft(3) = 7500 - 700 + 500
PLeft(4) = 8700 - 500 + 500
PLeft(5) = 9500 - 500 + 500
PLeft(6) = 3000 + 500
End Sub

Sub PrintTitle()
Printer.Orientation = 1
Printer.Font.Size = 11
Printer.Font.Name = "細明體"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 8000
Printer.CurrentX = 8000 + m_PrintLeftPos
Printer.CurrentY = 2700
Printer.Print Format(Now, "mmmm dd,YYYY")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 8000
Printer.CurrentX = 8000 + m_PrintLeftPos
Printer.CurrentY = 3000
Printer.Print "PAGE：" & str(Page)
iPrint = 3300
For i = 0 To 3
    If Len(strTemp(i)) <> 0 Then
        'Modify By Cheng 2002/12/30
'        Printer.CurrentX = 300
        Printer.CurrentX = 300 + m_PrintLeftPos
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
        iPrint = iPrint + 300
    End If
Next i
If Len(strTemp(4)) = 0 And Len(strTemp(5)) = 0 And Len(strTemp(6)) = 0 And Len(strTemp(7)) = 0 And Len(strTemp(8)) = 0 Then
    For i = 9 To 13
        If Len(strTemp(i)) <> 0 Then
            'Modify By Cheng 2002/12/30
'            Printer.CurrentX = 300
            Printer.CurrentX = 300 + m_PrintLeftPos
            Printer.CurrentY = iPrint
            Printer.Print strTemp(i)
            iPrint = iPrint + 300
        End If
    Next i
    'Add by Morgan 2011/5/26
    '地址6
    If Len(strTemp(26)) <> 0 Then
      Printer.CurrentX = 300 + m_PrintLeftPos
      Printer.CurrentY = iPrint
      Printer.Print strTemp(26)
      iPrint = iPrint + 300
    End If
Else
    For i = 4 To 8
        If Len(strTemp(i)) <> 0 Then
            'Modify By Cheng 2002/12/30
'            Printer.CurrentX = 300
            Printer.CurrentX = 300 + m_PrintLeftPos
            Printer.CurrentY = iPrint
            Printer.Print strTemp(i)
            iPrint = iPrint + 300
        End If
    Next i
End If
'Modify By Cheng 2002/12/20
'取消列印國家名稱
'Printer.CurrentX = 300
'Printer.CurrentY = iPrint
'Printer.Print strTemp(14)
iPrint = iPrint + 300
If Len(strTemp(15)) <> 0 Or Len(strTemp(16)) <> 0 Then
    'Modify By Cheng 2002/12/30
'    Printer.CurrentX = 300
    Printer.CurrentX = 300 + m_PrintLeftPos
    Printer.CurrentY = iPrint
    Printer.Print "ATTN："
    If Len(strTemp(15)) <> 0 Then
        'Modify By Cheng 2002/12/30
'        Printer.CurrentX = 1300
        Printer.CurrentX = 1300 + m_PrintLeftPos
        Printer.CurrentY = iPrint
        Printer.Print strTemp(15)
        iPrint = iPrint + 300
    End If
    If Len(strTemp(16)) <> 0 Then
        'Modify By Cheng 2002/12/30
'        Printer.CurrentX = 1300
        Printer.CurrentX = 1300 + m_PrintLeftPos
        Printer.CurrentY = iPrint
        Printer.Print strTemp(16)
        iPrint = iPrint + 300
    End If
End If
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Font.Name = "細明體"
Printer.Font.Size = 12
Printer.Print "Re："
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 1000
Printer.CurrentX = 1000 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "Reminder of the Annuities of Taiwan Patents"
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 1000
Printer.CurrentX = 1000 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "Due in " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY")
'Add By Cheng 2002/12/24
'年月加底線
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 1000 + Printer.TextWidth("Due in ")
Printer.CurrentX = 1000 + m_PrintLeftPos + Printer.TextWidth("Due in ")
Printer.CurrentY = iPrint + 300
Printer.DrawWidth = 2
Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.TextWidth(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY")), Printer.CurrentY)

iPrint = iPrint + 600
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
'Modified by Morgan 2024/4/10 對外統一用 Dear Colleagues --林總
'Printer.Print "Dear Sirs,"
Printer.Print "Dear Colleagues,"
'end 2024/4/10
iPrint = iPrint + 600
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/30
'Printer.Print "We would like to remind you that the annuity of the following Taiwan patents are to be paid "
Printer.Print "We would like to remind you that the annuity for the following Taiwan patents are "
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/30
'Printer.Print "before the due date in " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY")
Printer.Print "to be paid before the due date in " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 400 + Printer.TextWidth("before the due date in " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY"))
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 400 + m_PrintLeftPos + Printer.TextWidth("before the due date in " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY"))
Printer.CurrentX = 400 + m_PrintLeftPos + Printer.TextWidth("to be paid before the due date in " & Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY"))
Printer.CurrentY = iPrint
'Modify By Cheng 2002/12/30
'Printer.Print " and the estimated costs are listed for your reference："
Printer.Print " and the estimated costs are listed "
'Add By Cheng 2002/12/24
'年月加底線
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300 + Printer.TextWidth("before the due date in ")
Printer.CurrentX = 300 + m_PrintLeftPos + Printer.TextWidth("to be paid before the due date in ")
Printer.CurrentY = iPrint + 300
Printer.DrawWidth = 2
Printer.Line (Printer.CurrentX, Printer.CurrentY)-(Printer.CurrentX + Printer.TextWidth(Format(ChangeWStringToWDateString(ChangeTStringToWString(txt1(2))), "mmmm YYYY")), Printer.CurrentY)
'Add By Cheng 2002/12/30
iPrint = iPrint + 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print "for your reference："

Printer.Font.Size = 11
iPrint = iPrint + 300
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print String(150, "-")
iPrint = iPrint + 300
GetPleft
Printer.CurrentX = PLeft(0)
Printer.CurrentY = iPrint
Printer.Print "PAT NO."
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "APPLN NO."
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "YOUR REF"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "DUE DATE"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "YEAR"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "OUR REF"
iPrint = iPrint + 300
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "CASE NO."
iPrint = iPrint + 300
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 300
Printer.CurrentX = 300 + m_PrintLeftPos
Printer.CurrentY = iPrint
Printer.Print String(150, "-")
iPrint = iPrint + 300
End Sub
'Modify by Morgan 2005/4/11 加控制是否印欄位
'Sub PrintTitle1()
Sub PrintTitle1(Optional ByVal p_bolNoHeader As Boolean = False)
Printer.Font.Size = 11
Printer.Font.Name = "細明體"
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 8000
Printer.CurrentX = 8000 + m_PrintLeftPos
Printer.CurrentY = 2700
Printer.Print Format(Now, "mmmm dd,YYYY")
'Modify By Cheng 2002/12/30
'Printer.CurrentX = 8000
Printer.CurrentX = 8000 + m_PrintLeftPos
Printer.CurrentY = 3000
Printer.Print "PAGE：" & str(Page)
iPrint = 3300
   'Add by Morgan 2005/4/11 加判斷是否印欄位
   If p_bolNoHeader = False Then
      'Modify By Cheng 2002/12/30
      'Printer.CurrentX = 300
      Printer.CurrentX = 300 + m_PrintLeftPos
      Printer.CurrentY = iPrint
      Printer.Print String(150, "-")
      iPrint = iPrint + 300
      GetPleft
      Printer.CurrentX = PLeft(0)
      Printer.CurrentY = iPrint
      Printer.Print "PAT NO."
      Printer.CurrentX = PLeft(1)
      Printer.CurrentY = iPrint
      Printer.Print "APPLN NO."
      Printer.CurrentX = PLeft(2)
      Printer.CurrentY = iPrint
      Printer.Print "YOUR REF"
      Printer.CurrentX = PLeft(3)
      Printer.CurrentY = iPrint
      Printer.Print "DUE DATE"
      Printer.CurrentX = PLeft(4)
      Printer.CurrentY = iPrint
      Printer.Print "YEAR"
      Printer.CurrentX = PLeft(5)
      Printer.CurrentY = iPrint
      Printer.Print "OUR REF"
      iPrint = iPrint + 300
      Printer.CurrentX = PLeft(6)
      Printer.CurrentY = iPrint
      Printer.Print "CASE NO."
      iPrint = iPrint + 300
      'Modify By Cheng 2002/12/30
      'Printer.CurrentX = 300
      Printer.CurrentX = 300 + m_PrintLeftPos
      Printer.CurrentY = iPrint
      Printer.Print String(150, "-")
      iPrint = iPrint + 300
   End If

End Sub

Private Sub Form_Load()
'Add By Cheng 2003/02/05
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim ii As Integer

MoveFormToCenter Me
strSql = "SELECT USXR01,USXR02 FROM USXRATE order by usxr01 desc "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
strSql = ""
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    strSql = CheckStr(adoRecordset.Fields(1))
End If
CheckOC
StrTemp99(0) = "1600"
'Modify By Cheng 2002/12/20
'StrTemp99(1) = "1800"
StrTemp99(1) = "2500"
StrTemp99(2) = CheckStr(Val(StrTemp99(0)) + Val(StrTemp99(1)))
StrTemp99(3) = ""
StrTemp99(4) = "1600"
'Modify By Cheng 2002/12/20
'StrTemp99(5) = "3600"
StrTemp99(5) = "5000"
StrTemp99(6) = CheckStr(Val(StrTemp99(4)) + Val(StrTemp99(5)))
StrTemp99(7) = ""
StrTemp99(8) = "1600"
'Modify By Cheng 2002/12/20
'StrTemp99(9) = "7200"
'Modify by Morgan 2004/8/10
'StrTemp99(9) = "10000"
StrTemp99(9) = "9000"
StrTemp99(10) = CheckStr(Val(StrTemp99(8)) + Val(StrTemp99(9)))
StrTemp99(11) = ""
StrTemp99(12) = "1600"
'Modify By Cheng 2002/12/20
'StrTemp99(13) = "14400"
'Modify by Morgan 2004/8/10
'StrTemp99(13) = "20000"
StrTemp99(13) = "18000"
StrTemp99(14) = CheckStr(Val(StrTemp99(12)) + Val(StrTemp99(13)))
StrTemp99(15) = ""
If Val(strSql) <> 0 Then
    StrTemp99(3) = CheckStr(Val(StrTemp99(2)) / Val(strSql))
    StrTemp99(7) = CheckStr(Val(StrTemp99(6)) / Val(strSql))
    StrTemp99(11) = CheckStr(Val(StrTemp99(10)) / Val(strSql))
    StrTemp99(15) = CheckStr(Val(StrTemp99(14)) / Val(strSql))
End If

'Modify by Morgan 2011/3/15 改共用且不要排除預設印表機
   PUB_SetPrinter Me.Name, Combo1, strPrinter
'end 2011/3/15
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Morgan 2010/3/11
   If m_AddrList <> "" Then
      If MsgBox("準備列印地址條，請更換紙張!!!", vbExclamation + vbOKCancel) = vbOK Then
         Load frm083014: DoEvents
         With frm083014
            '隱藏表單
            .Hide
            .opt1(1).Value = True
            .Text1(1).Text = m_AddrList
            '設定印表機
            .SetPrinter Combo1.Text
            '列印份數
            .Text1(3).Text = "1"
            '是否含不寄雜誌的對象
            .Text1(5).Text = "Y"
            .Text1(4).Text = "2"
            '執行列印
            .cmdPrint_Click: DoEvents
            .cmdBack.Value = True
         End With
      End If
   End If

   'Copy from cmdok_Click by Morgan 2004/10/26
   '列印地址條
   PUB_PrintAddressList strUserNum, Me.Combo1.Text
   '刪除地址條列表資料
   PUB_DeleteAddressList strUserNum
   '初始化序號
   pub_AddressListSN = 0
   '若印表機變動, 則更新列印設定
   If Me.Combo1.Text <> Me.Combo1.Tag Then
       PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2004/10/26 end
   Set frm060305 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
txt1(Index).SelStart = 0
txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    cmdOK(0).SetFocus
End If
End Sub
Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0 '代理人(申請人)
    'Add By Cheng 2003/02/17
    '若未輸入資料則不檢查
    If Me.txt1(Index).Text = "" Then Exit Sub
    If Mid(UCase(txt1(0)), 1, 1) = "X" Then
        strSql = "select nvl(cu05||cu88||cu89||cu90,nvl(cu04,cu06)) from customer where cu01='" & Mid(GetNewFagent(txt1(0)), 1, 8) & "' and cu02='" & Mid(GetNewFagent(txt1(0)), 9, 1) & "' "
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            lbl1.Caption = CheckStr(adoRecordset.Fields(0))
        Else
            MsgBox "代理人(申請人)代號錯誤，請重新輸入 !", vbCritical
            lbl1.Caption = ""
            txt1(0).SetFocus
        End If
        CheckOC
    ElseIf Mid(UCase(txt1(0)), 1, 1) = "Y" Then
        strSql = "select nvl(fa05||fa63||fa64||fa65,nvl(fa04,fa06)) from fagent where fa01='" & Mid(GetNewFagent(txt1(0)), 1, 8) & "' and fa02='" & Mid(GetNewFagent(txt1(0)), 9, 1) & "' "
        CheckOC
        adoRecordset.CursorLocation = adUseClient
        adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
            lbl1.Caption = CheckStr(adoRecordset.Fields(0))
        Else
            MsgBox "代理人(申請人)代號錯誤，請重新輸入 !", vbCritical
            lbl1.Caption = ""
            txt1(0).SetFocus
        End If
        CheckOC
    Else
        MsgBox "代理人(申請人)代號錯誤，請重新輸入 !", vbCritical
        lbl1.Caption = ""
        txt1(0).SetFocus
    End If
Case 2
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
     If RunNick(txt1(1), txt1(2)) Then
         txt1(1).SetFocus
         txt1_GotFocus (1)
     End If
   Else
      blnClkSure = False
   End If
Case Else
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   If txt1(Index) = "" Then Exit Sub
   Select Case Index
      Case 1, 2 '年費期限
         Cancel = Not ChkDate(txt1(Index))
         If Cancel Then TextInverse txt1(Index)
   End Select
End Sub

'Add By Cheng 2002/12/30
'改變年的列印格式
Private Function ChgYearFormat(strYear As String) As String
    ChgYearFormat = ""
    Select Case strYear
    Case "1"
        ChgYearFormat = "" & strYear & "st"
    Case "2"
        ChgYearFormat = "" & strYear & "nd"
    Case "3"
        ChgYearFormat = "" & strYear & "rd"
    Case Else
        ChgYearFormat = "" & strYear & "th"
    End Select
End Function
