VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090706 
   BorderStyle     =   1  '單線固定
   Caption         =   "繪圖人員工作進度資料查詢"
   ClientHeight    =   2580
   ClientLeft      =   3090
   ClientTop       =   2595
   ClientWidth     =   4275
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   4275
   Begin VB.Frame Frame1 
      Height          =   2025
      Left            =   60
      TabIndex        =   9
      Top             =   420
      Width           =   4185
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   3
         Left            =   1050
         MaxLength       =   3
         TabIndex        =   4
         Top             =   1080
         Width           =   480
      End
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   4
         Left            =   1980
         MaxLength       =   2
         TabIndex        =   5
         Top             =   1080
         Width           =   480
      End
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   0
         Left            =   1050
         MaxLength       =   1
         TabIndex        =   1
         Top             =   180
         Width           =   270
      End
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   1
         Left            =   1500
         MaxLength       =   1
         TabIndex        =   2
         Top             =   180
         Width           =   270
      End
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   2
         Left            =   1050
         MaxLength       =   6
         TabIndex        =   3
         Top             =   630
         Width           =   900
      End
      Begin VB.TextBox Txt1 
         Height          =   300
         Index           =   5
         Left            =   1035
         MaxLength       =   1
         TabIndex        =   6
         Top             =   1530
         Width           =   315
      End
      Begin MSForms.Label lbl1 
         Height          =   255
         Index           =   0
         Left            =   2010
         TabIndex        =   19
         Top             =   660
         Width           =   1965
         Caption         =   "lblFM2"
         Size            =   "3466;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "發文年月："
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   17
         Top             =   1110
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "年"
         Height          =   180
         Index           =   2
         Left            =   1635
         TabIndex        =   16
         Top             =   1110
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "月"
         Height          =   180
         Index           =   7
         Left            =   2595
         TabIndex        =   15
         Top             =   1110
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "所別："
         Height          =   180
         Index           =   4
         Left            =   60
         TabIndex        =   14
         Top             =   240
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "繪圖人員："
         Height          =   180
         Index           =   8
         Left            =   60
         TabIndex        =   13
         Top             =   675
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "(1.螢幕 2.報表)"
         Height          =   180
         Index           =   21
         Left            =   1440
         TabIndex        =   12
         Top             =   1560
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "顯示方式："
         Height          =   180
         Index           =   22
         Left            =   60
         TabIndex        =   11
         Top             =   1560
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "(1.北 2.中 3.南 4.高 5.其他)"
         Height          =   180
         Index           =   24
         Left            =   1830
         TabIndex        =   10
         Top             =   210
         Width           =   2160
      End
      Begin VB.Line Line3 
         X1              =   1170
         X2              =   1620
         Y1              =   345
         Y2              =   345
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Index           =   1
      Left            =   3012
      TabIndex        =   8
      Top             =   20
      Width           =   1200
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   2232
      TabIndex        =   7
      Top             =   20
      Width           =   756
   End
   Begin VB.TextBox Txt1 
      Height          =   264
      Index           =   6
      Left            =   1710
      MaxLength       =   1
      TabIndex        =   0
      Top             =   450
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "是否含已發文資料：                (Y/N)"
      Height          =   195
      Left            =   60
      TabIndex        =   18
      Top             =   480
      Width           =   2865
   End
End
Attribute VB_Name = "frm090706"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/01/28 改成Form2.0 ; lbl1(index); Printer列印未改
'Memo By Morgan 2012/12/10 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/17 日期欄已修改
Option Explicit

Public TextOk As Boolean
Dim strSql As String, strSQL1 As String, i As Integer, j As Integer, s As Integer, SavDay1 As String, SavDay2 As String, SavDay3 As String, StrSQL7 As String, StrSQL4 As String, strSQL5 As String
'Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 22) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String, K As Integer
Dim strSQL2 As String, iPrint As Integer, Page As Integer, strTemp(0 To 24) As String, strTemp3 As String, TestOk As Boolean, StrTemp99(0 To 21) As String, k As Integer
'Dim Pleft(0 To 22) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, TempSeekNick As String
Dim PLeft(0 To 24) As Integer, strTemp1 As Variant, strTemp2 As Variant, StrSQL6 As String, Str020401SysKind As String, Seekok As Integer, SeekTemp As Integer, TempSeekNick As String
Dim strDate1 As String, StrDate2 As String
Dim NickRS As ADODB.Recordset
Dim strSQL8 As String, strSQL9 As String
Dim m_ProState As String 'Add By Sindy 2017/8/10 記錄目前權限


Private Sub cmdOK_Click(Index As Integer)
Select Case Index
Case 0 '確定
     If Len(txt1(3)) = 0 Then
         s = MsgBox("發文年不可空白!!", , "USER 輸入錯誤")
         txt1(3).SetFocus
         Exit Sub
     Else
         If Len(txt1(4)) = 0 Then
             s = MsgBox("發文月不可空白!!", , "USER 輸入錯誤")
             txt1(4).SetFocus
             Exit Sub
         Else
             If Len(txt1(5)) = 0 Then
                 s = MsgBox("顯示方式不可空白!!", , "USER 輸入錯誤")
                 txt1(5).SetFocus
                 Exit Sub
             Else
                 '查詢
                 If Trim(txt1(5)) = "1" Then
                    Screen.MousePointer = vbHourglass
                    Me.Enabled = False
                    frm090711.Show
                    If TextOk = False Then
                        Unload frm090711
                        Me.Show
                    Else
                        Me.Hide
                    End If
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                 '印表
                 Else
                    Screen.MousePointer = vbHourglass
                    Me.Enabled = False
                    'add by nick 2005/01/12 統計改用查詢那支的程式
                    Load frm090711
                    frm090711.Hide
                    'add by nickc 2005/04/13 算所有繪圖的統計
                    frm090711.From090706 = True
                    frm090711.StrMenu2
                    'add end
                    'edit by nickc 2005/04/13 因為 frm090711 已經有資料，所以不用重抓，直接用 frm090711 的資料
                    Process
                    Process1
                    'add by nick 2005/01/12 統計改用查詢那支的程式
                    Unload frm090711
                    'add end
                    Me.Enabled = True
                    Screen.MousePointer = vbDefault
                End If
             End If
         End If
     End If
Case 1 '回前畫面
     Unload Me
Case Else
End Select
End Sub

Sub Process2()
ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/16 清除查詢印表記錄檔欄位
pub_QL05 = pub_QL05 & ";" & Label1(22) & "1.螢幕" 'Add By Sindy 2010/12/16
cnnConnection.Execute "DELETE FROM R090706 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090706_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090706_2 WHERE ID='" & strUserNum & "' "
StrSQL6 = ""
If Len(txt1(0)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ST06>='" & txt1(0) & "' "
End If
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND sT06<='" & txt1(1) & "' "
End If
If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(0) & "-" & txt1(1) & Label1(24) 'Add By Sindy 2010/12/16
End If
If Len(txt1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND st01='" & txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(2) & lbl1(0) 'Add By Sindy 2010/12/16
End If
If txt1(3) <> "" Or txt1(4) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(3) & txt1(4) 'Add By Sindy 2010/12/16
End If
frm090711.Show
frm090711.Hide
CheckOC
With adoRecordset
       .CursorLocation = adUseClient
        'Modify By Cheng 2003/09/22
        'Begin
'       .Open "select DISTINCT st01 from staff where st04='1' AND ST05 in ('79','81','82','AC') " & StrSQL6, cnnConnection, adOpenStatic, adLockReadOnly
'       .Open "select DISTINCT st01||' '||ST02 from staff where st04='1' AND ST05 in ('79','81','82','AC') " & StrSQL6, cnnConnection, adOpenStatic, adLockReadOnly
        '2011/9/20 MODIFY BY SONIA 指定人員時不判斷是否在職93011
        '.Open "Select ST01||' '||ST02 From Staff Where ST04='1' AND ST05 In ('79','81','82','AC') " & StrSQL6 & " Order By ST06, ST01 ", cnnConnection, adOpenStatic, adLockReadOnly
        If Len(txt1(2)) <> 0 Then
           .Open "Select ST01||' '||ST02 From Staff Where ST05 In ('79','81','82','AC') " & StrSQL6 & " Order By ST06, ST01 ", cnnConnection, adOpenStatic, adLockReadOnly
        Else
           .Open "Select ST01||' '||ST02 From Staff Where ST04='1' AND ST05 In ('79','81','82','AC') " & StrSQL6 & " Order By ST06, ST01 ", cnnConnection, adOpenStatic, adLockReadOnly
        End If
        '2011/9/20 END
        'End
       If .RecordCount <> 0 Then
            frm090711.Combo1.Clear
            s = 0
            Do While .EOF = False
                frm090711.Combo1.AddItem CheckStr(.Fields(0)), s
                s = s + 1
                .MoveNext
            Loop
            frm090711.Combo1.Text = frm090711.Combo1.List(0)
            TextOk = True
        Else
            TextOk = False
        End If
End With

End Sub

Sub Process()
Dim strMonthLastDate As String
Dim strBeginDate As String
Dim strEndDate As String

ClearQueryLog (Me.Name) 'Add By Sindy 2010/12/16 清除查詢印表記錄檔欄位
pub_QL05 = pub_QL05 & ";" & Label1(22) & "2.報表" 'Add By Sindy 2010/12/16
cnnConnection.Execute "DELETE FROM R090706 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090706_1 WHERE ID='" & strUserNum & "' "
cnnConnection.Execute "DELETE FROM R090706_2 WHERE ID='" & strUserNum & "' "
StrSQL6 = "": strSQL1 = "": StrSQL7 = "": strSQL8 = "": strSQL9 = ""
'所別(起)
If Len(txt1(0)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND s2.ST06>='" & txt1(0) & "' "
    strSQL1 = strSQL1 + " AND s2.ST06>='" & txt1(0) & "' "
    StrSQL7 = StrSQL7 + " AND s2.ST06>='" & txt1(0) & "' "
    strSQL8 = strSQL8 + " AND s2.ST06>='" & txt1(0) & "' "
    strSQL9 = strSQL9 + " AND s2.ST06>='" & txt1(0) & "' "
End If
'所別(迄)
If Len(txt1(1)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND s2.sT06<='" & txt1(1) & "' "
    strSQL1 = strSQL1 + " AND s2.sT06<='" & txt1(1) & "' "
    StrSQL7 = StrSQL7 + " AND s2.sT06<='" & txt1(1) & "' "
    strSQL8 = strSQL8 + " AND s2.sT06<='" & txt1(1) & "' "
    strSQL9 = strSQL9 + " AND s2.sT06<='" & txt1(1) & "' "
End If
If Len(txt1(0)) <> 0 Or Len(txt1(1)) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label1(4) & txt1(0) & "-" & txt1(1) & Label1(24) 'Add By Sindy 2010/12/16
End If
'繪圖人員
If Len(txt1(2)) <> 0 Then
    StrSQL6 = StrSQL6 + " AND ep13='" & txt1(2) & "' "
    strSQL1 = strSQL1 + " AND ep13='" & txt1(2) & "' "
    StrSQL7 = StrSQL7 + " AND ep13='" & txt1(2) & "' "
    strSQL8 = strSQL8 + " AND ep13='" & txt1(2) & "' "
    strSQL9 = strSQL9 + " AND SH02='" & txt1(2) & "' "
    pub_QL05 = pub_QL05 & ";" & Label1(8) & txt1(2) & lbl1(0) 'Add By Sindy 2010/12/16
End If

If txt1(3) <> "" Or txt1(4) <> "" Then
   pub_QL05 = pub_QL05 & ";" & Label1(0) & txt1(3) & txt1(4) 'Add By Sindy 2010/12/16
End If

'Modify By Cheng 2003/05/28
StrSQL6 = StrSQL6 & " And CP01 in ('FCP','P','CFP')  "
StrSQL7 = StrSQL7 & " And CP01 in ('FCP','P','CFP')  "
strSQL8 = strSQL8 & " And CP01 in ('FCP','P','CFP')  "
strSQL9 = strSQL9 & " And SH06 in ('FCP','P','CFP')  "
      
'若進入個人工作維護
If ProState = "1" Then
    StrSQL6 = StrSQL6 & "  AND EP13='" & strUserNum & "' "
    StrSQL7 = StrSQL7 & "  AND EP13='" & strUserNum & "' "
    strSQL8 = strSQL8 & "  AND EP13='" & strUserNum & "' "
    '若不含已發文資料
    If Me.txt1(6).Text <> "Y" Then
        StrSQL6 = StrSQL6 & "  AND CP27 Is Null "
    End If
    strSQL9 = strSQL9 & "  AND SH02='" & strUserNum & "' "
End If
'未發文, 未取消收文
If ProState = 2 Then
    'Modify By Cheng 2003/05/28
'   StrSQL6 = StrSQL6 & " AND CP05<=" & Val(Me.Txt1(3).Text & Me.Txt1(4).Text) + 191100 & "31 "
   StrSQL6 = StrSQL6 & " AND CP05<=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "31 "
   StrSQL7 = StrSQL7 & " AND CP05<=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "31 "
End If
'Modify By Cheng 2003/05/28
'strSQL6 = strSQL6 + " and CP26 IS NULL  and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
StrSQL6 = StrSQL6 + " and CP27 IS NULL  and CP57 IS NULL and cp05>=19980101 "
StrSQL7 = StrSQL7 + " And cp05>=19980101 "

'strSQL1 = " and ep20 is null  AND EP13='" & strUserNum & "' and cp01 in ('FCP','P','CFP')  "
strSQL1 = strSQL1 & " and cp01 in ('FCP','P','CFP')  "
If ProState = "1" Then
    strSQL1 = strSQL1 & " AND EP13='" & strUserNum & "' "
    '若不含已發文資料
    If Me.txt1(6).Text <> "Y" Then
        strSQL1 = strSQL1 & "  AND CP27 Is Null "
    End If
    strSQL9 = strSQL9 & " AND SH02='" & strUserNum & "' "
End If
'Modify By Cheng 2002/04/22
'(未發文, 當月取消收文) 或 (當月發文, 未取消收文)
'strSQL1 = strSQL1 & " and CP26 IS NULL  and (SUBSTR(CP27,1,6)=" & Mid(GetTodayDate, 1, 6) & " or SUBSTR(CP57,1,6)=" & Mid(GetTodayDate, 1, 6) & ") and cp05>=19980101 "
'若為繪圖人員個人作業時
If ProState = 1 Then
    'Modify By Cheng 2003/05/28
'   strSQL1 = strSQL1 & " and CP26 IS NULL and ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & " AND CP57 IS NULL) " & " or ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " )) and cp05>=19980101 "
   'edit by nickc 2005/05/12
   'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & Mid(strSrvDate(1), 1, 6) & " AND CP57 IS NULL) " & " or ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & Mid(strSrvDate(1), 1, 6) & " )) and cp05>=19980101 "
   'edit by nickc 2005/05/13
   'strSQL1 = strSQL1 & " and ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 AND CP57 IS NULL) " & " or ( CP27 IS NULL AND CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp57<=" & Mid(strSrvDate(1), 1, 6) & "31 )) and cp05>=19980101 "
   strSQL1 = strSQL1 & " and ((CP27>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp27<=" & Mid(strSrvDate(1), 1, 6) & "31 ) " & " or ( CP27 IS NULL AND CP57>=" & Mid(strSrvDate(1), 1, 6) & "01 and cp57<=" & Mid(strSrvDate(1), 1, 6) & "31 )) and cp05>=19980101 "
'若為繪圖人員管理作業時
ElseIf ProState = 2 Then
    'Modify By Cheng 2003/05/28
'   strSQL1 = strSQL1 & " AND CP05<=" & Val(Me.txt1(3).Text & Me.txt1(4).Text) + 191100 & "31 "
'   strSQL1 = strSQL1 & " AND CP05<=" & Val(Me.txt1(3).Text & Me.txt1(4).Text & "31") + 19110000
'   strSQL1 = strSQL1 & " and CP26 IS NULL and ((SUBSTR(CP27,1,6)=" & Val(Me.Txt1(3).Text & Me.Txt1(4).Text) + 191100 & " AND CP57 IS NULL) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & Val(Me.Txt1(3).Text & Val(Me.Txt1(4).Text)) + 191100 & " )) and cp05>=19980101 "
   strSQL1 = strSQL1 & " AND CP05<=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00") & "31") + 19110000
    'Modify By Cheng 2003/05/28
'   strSQL1 = strSQL1 & " and CP26 IS NULL and ((SUBSTR(CP27,1,6)=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & " AND CP57 IS NULL) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & " )) and cp05>=19980101 "
   'edit by nickc 2005/05/12
   'strSQL1 = strSQL1 & " and ((SUBSTR(CP27,1,6)=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & " AND CP57 IS NULL) OR ( CP27 IS NULL AND SUBSTR(CP57,1,6)=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & " )) and cp05>=19980101 "
   'edit by nickc 2005/05/13
   'strSQL1 = strSQL1 & " and ((CP27>=" & Val(Format(Me.Txt1(3).Text, "000") & Format(Me.Txt1(4).Text, "00")) + 191100 & "01 and cp27<=" & Val(Format(Me.Txt1(3).Text, "000") & Format(Me.Txt1(4).Text, "00")) + 191100 & "31 AND CP57 IS NULL) OR ( CP27 IS NULL AND CP57>=" & Val(Format(Me.Txt1(3).Text, "000") & Format(Me.Txt1(4).Text, "00")) + 191100 & "01 and cp57<=" & Val(Format(Me.Txt1(3).Text, "000") & Format(Me.Txt1(4).Text, "00")) + 191100 & "31 )) and cp05>=19980101 "
   strSQL1 = strSQL1 & " and ((CP27>=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "01 and cp27<=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "31 ) OR ( CP27 IS NULL AND CP57>=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "01 and cp57<=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "31 )) and cp05>=19980101 "
End If
'edit by nick 2005/03/01 墨圖也要判斷
'StrSQL6 = StrSQL6 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
'strSQL1 = strSQL1 & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
StrSQL6 = StrSQL6 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
strSQL1 = strSQL1 & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
strSQL9 = strSQL9 & " And SUBSTR(SH01,1,6)=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & " "
      'add by nickc 2005/03/21
      strSQL1 = strSQL1 & " and  cp107='Y' "
      StrSQL6 = StrSQL6 & " and cp107='Y' "
CheckOC

'add by nickc 2005/05/12 CFP 草不計件的，墨未完稿的不印
strSQL1 = strSQL1 & " and not (cp01='CFP' and ep20 is not null and ep20='N' and ep18 is null) "
StrSQL6 = StrSQL6 & " and not (cp01='CFP' and ep20 is not null and ep20='N' and ep18 is null) "

'add by nickc 2006/06/30 跟螢幕資料同步
strSQL1 = strSQL1 & " and ((pa58>=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "01 and pa58<=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "31) or pa58 is null) "
StrSQL6 = StrSQL6 & " and ((pa58>=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "01 and pa58<=" & Val(Format(Me.txt1(3).Text, "000") & Format(Me.txt1(4).Text, "00")) + 191100 & "31) or pa58 is null) "

'Modify By Cheng 2002/04/17
'strSQL = "SELECT nvl(S2.ST02,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL & " UNION all  SELECT nvl(S2.ST02,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " ORDER BY 1 "
'92.04.03 nick add left join
'strSQL = "SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL & " UNION all  SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND PA01=CP01 AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'strSQL = strSQL + " ORDER BY 1 "
'Modify By Cheng 2003/06/30
'strSQL = "SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & StrSQL6
'strSQL = strSQL & " UNION all  SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) " & strSQL1
'edit by nick 2004/10/26
'strSQL = "SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19, EP29, Nvl(EP06, 0), EP26  FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 Is Not Null " & StrSQL6
'strSQL = strSQL & " UNION all  SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04,NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19, EP29, Nvl(EP06, 0), EP26 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 IS Not Null " & strSQL1
'edit by nickc 加計件值
'strSQL = "SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19, EP29, Nvl(EP06, 0), EP26  FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 Is Not Null " & StrSQL6
'strSQL = strSQL & " UNION all  SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),eP20,decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19, EP29, Nvl(EP06, 0), EP26 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 IS Not Null " & strSQL1
strSql = "SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),nvl(EP20,round(cp100 * cp101,2)),decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19, nvl(EP29,round(cp103 * cp104,2)), Nvl(EP06, 0), EP26  FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 Is Not Null " & StrSQL6
strSql = strSql & " UNION all  SELECT nvl(S2.ST01,ep13),SUBSTR(CP09,1,1)," & SQLDate("CP05") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04||DECODE(PA57,'Y','＊',''),NVL(PA05,NVL(PA06,PA07)),nvl(EP20,round(cp100 * cp101,2)),decode(pa09,'000',cpm03,cpm04),s1.st02," & SQLDate("CP48") & ",cp18," & SQLDate("eP14") & "," & SQLDate("eP15") & ",0," & SQLDate("EP17") & "," & SQLDate("EP18") & ",0," & SQLDate("cP06") & "," & SQLDate("cP27") & ",ep26,s3.st02,CP09,ep21,ep22,ep23,ep24,ep25,DECODE(PA09,'000',PTM03,PTM04),ep16,ep19, nvl(EP29,round(cp103 * cp104,2)), Nvl(EP06, 0), EP26 FROM CASEPROGRESS,ENGINEERPROGRESS,STAFF S1,STAFF S2,STAFF S3,PATENT,CASEPROPERTYMAP,PATENTTRADEMARKMAP WHERE EP02=CP09(+) AND cp01=pa01(+) and CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP05=S1.ST01(+) AND eP13=S2.ST01(+) AND cP13=S3.ST01(+) and cp01=cpm01(+) and cp10=cpm02(+) AND '1'=PTM01(+)  AND PA08=PTM02(+) And EP13 IS Not Null " & strSQL1

strSql = strSql + " ORDER BY 1 "

'CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        InsertQueryLog (.RecordCount) 'Add By Sindy 2010/12/16
        .MoveFirst
        DoEvents
        Do While .EOF = False
            For i = 0 To 19
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            '計算草圖作業天數
            'If Len(strTemp(11)) <> 0 And Len(strTemp(10)) <> 0 And Val(strTemp(10)) <> 0 And Val(strTemp(11)) <> 0 Then
            '    strTemp(12) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(11))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(10))))))
            'Else
            '    strTemp(12) = "0"
            'End If
            strDate1 = CheckStr(.Fields(11))
            StrDate2 = CheckStr(.Fields(10))
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
                strTemp(12) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
                strTemp(12) = "0"
            End If
            '計算墨圖作業天數
            strDate1 = CheckStr(.Fields(14))
            StrDate2 = CheckStr(.Fields(13))
            If Trim(strDate1) <> "" And Trim(StrDate2) <> "" Then
'                strTemp(14) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(StrDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
                strTemp(15) = GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strDate1)), ChangeTStringToWString(ChangeTDateStringToTString(StrDate2)))
            Else
'                strTemp(14) = "0"
                strTemp(15) = "0"
            End If
            'If Len(strTemp(14)) <> 0 And Len(strTemp(13)) <> 0 And Val(strTemp(14)) <> 0 And Val(strTemp(13)) <> 0 Then
            '     strTemp(15) = Trim(str(GetWorkDay(ChangeTStringToWString(ChangeTDateStringToTString(strTemp(14))), ChangeTStringToWString(ChangeTDateStringToTString(strTemp(13))))))
            'Else
            '    strTemp(15) = "0"
            'End If
            strSql = "INSERT INTO R090706 VALUES ('" & strTemp(0) & "','" & strTemp(1) & "','" & strTemp(2) & "','" & strTemp(3) & "','" & ChgSQL(strTemp(4)) & "','" & strTemp(5) & "','" & strTemp(6) & "','" & strTemp(7) & "','" & strTemp(8) & "'," & Val(strTemp(9)) & ",'" & strTemp(10) & "','" & strTemp(11) & "'," & Val(strTemp(12)) & ",'" & strTemp(13) & "','" & strTemp(14) & "'," & Val(strTemp(15)) & ",'" & strTemp(16) & "','" & strTemp(17) & "','" & strTemp(18) & "','" & strTemp(19) & "','" & CheckStr(.Fields(20)) & "','" & strUserNum & "') "
            cnnConnection.Execute strSql
            TempSeekNick = ""
            'Modify By Cheng 2003/07/01
'            '草圖張數 + 墨圖張數
'            Select Case Val(CheckStr(.Fields(27))) + Val(CheckStr(.Fields(28)))
''            Case Is <= 21, Is >= 15
''                 TempSeekNick = "2"
''            Case Is >= 22, Is <= 28
''                 TempSeekNick = "3"
''            Case Is >= 29
''                 TempSeekNick = "4"
'            Case 15 To 21
'                 TempSeekNick = "2"
'            Case 22 To 28
'                 TempSeekNick = "3"
'            Case Is >= 29
'                 TempSeekNick = "4"
'            Case Else
'                 TempSeekNick = "0"
'            End Select
            If InStr("" & .Fields(31).Value, "複雜") > 0 Then
                TempSeekNick = "1"
            End If
            'Modify By Cheng 2002/04/17
'            strSQL = "insert into r090706_2 values ('" & strTemp(0) & "','" & strTemp(1) & "','" & strTemp(2) & "','" & CheckStr(.Fields(26)) & "'," & Val(strTemp(9)) & ",'" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(7) & "','" & strTemp(10) & "','" & strTemp(11) & "'," & Val(CheckStr(.Fields(27))) & "," & Val(strTemp(12)) & ",'" & strTemp(13) & "','" & strTemp(14) & "'," & Val(CheckStr(.Fields(28))) & "," & Val(strTemp(15)) & "," & Val(CheckStr(.Fields(21))) & "," & Val(CheckStr(.Fields(22))) & "," & Val(TempSeekNick) & "," & Val(CheckStr(.Fields(23))) & "," & Val(CheckStr(.Fields(24))) & "," & Val(CheckStr(.Fields(25))) & ",'" & strTemp(19) & "','" & strUserNum & "') "
            strTemp(18) = LeftB("" & strTemp(18), 20)
            'Modify By Cheng 2003/06/30
'            strSQL = "insert into r090706_2 values ('" & strTemp(0) & "','" & strTemp(1) & "','" & strTemp(2) & "','" & CheckStr(.Fields(26)) & "'," & Val(strTemp(9)) & ",'" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(7) & "','" & strTemp(10) & "','" & strTemp(11) & "'," & Val(CheckStr(.Fields(27))) & "," & Val(strTemp(12)) & ",'" & strTemp(13) & "','" & strTemp(14) & "'," & Val(CheckStr(.Fields(28))) & "," & Val(strTemp(15)) & "," & Val(CheckStr(.Fields(21))) & "," & Val(CheckStr(.Fields(22))) & "," & Val(TempSeekNick) & "," & Val(CheckStr(.Fields(23))) & "," & Val(CheckStr(.Fields(24))) & "," & Val(CheckStr(.Fields(25))) & ",'" & strTemp(18) & "','" & strUserNum & "','" & strTemp(5) & "') "
            'edit by nickc 2005/05/12 收文日位置改放發文日
            'strSQL = "insert into r090706_2 values ('" & strTemp(0) & "','" & strTemp(1) & "','" & strTemp(2) & "','" & CheckStr(.Fields(26)) & "'," & Val(strTemp(9)) & ",'" & strTemp(3) & "','" & strTemp(4) & "','" & strTemp(7) & "','" & strTemp(10) & "','" & strTemp(11) & "'," & Val(CheckStr(.Fields(27))) & "," & Val(strTemp(12)) & ",'" & strTemp(13) & "','" & strTemp(14) & "'," & Val(CheckStr(.Fields(28))) & "," & Val(strTemp(15)) & "," & Val(CheckStr(.Fields(21))) & "," & Val(CheckStr(.Fields(22))) & "," & Val(TempSeekNick) & "," & Val(CheckStr(.Fields(23))) & "," & Val(CheckStr(.Fields(24))) & "," & Val(CheckStr(.Fields(25))) & ",'" & strTemp(18) & "','" & strUserNum & "','" & strTemp(5) & "','" & .Fields(29).Value & "'," & .Fields(30).Value & ") "
            strSql = "insert into r090706_2 values ('" & strTemp(0) & "','" & strTemp(1) & "','" & strTemp(17) & "','" & CheckStr(.Fields(26)) & "'," & Val(strTemp(9)) & ",'" & strTemp(3) & "','" & ChgSQL(strTemp(4)) & "','" & strTemp(7) & "','" & strTemp(10) & "','" & strTemp(11) & "'," & Val(CheckStr(.Fields(27))) & "," & Val(strTemp(12)) & ",'" & strTemp(13) & "','" & strTemp(14) & "'," & Val(CheckStr(.Fields(28))) & "," & Val(strTemp(15)) & "," & Val(CheckStr(.Fields(21))) & "," & Val(CheckStr(.Fields(22))) & "," & Val(TempSeekNick) & "," & Val(CheckStr(.Fields(23))) & "," & Val(CheckStr(.Fields(24))) & "," & Val(CheckStr(.Fields(25))) & ",'" & strTemp(18) & "','" & strUserNum & "','" & strTemp(5) & "','" & .Fields(29).Value & "'," & .Fields(30).Value & ") "
            cnnConnection.Execute strSql
            .MoveNext
            DoEvents
        Loop
        TextOk = True
    Else
        InsertQueryLog (0) 'Add By Sindy 2010/12/16
        ShowNoData
        TextOk = False
        Exit Sub
    End If
    CheckOC
    
    strMonthLastDate = (Val(Me.txt1(3).Text) + 1911) & Format(Me.txt1(4).Text, "00") & PUB_GetMonthDays(Val(Me.txt1(3).Text) + 1911, Val(Me.txt1(4).Text))
    '統計其他項目
    '可辦草圖
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) select EP13,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND cp01 in ('CFP','P') AND EP13='" & Trim(txt1(2).Text) & "' AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL GROUP BY EP13 "
'    strSQL = "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) select EP13,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND cp01 in ('CFP','P','FCP') AND EP13='" & Trim(txt1(2).Text) & "' AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL And EP29 Is Null GROUP BY EP13 "
    strSql = ""
    If ProState = "2" Or ProState = "3" Then '管理
        '若發文年月<系統年月
        If Val(Me.txt1(3).Text) + 1911 & Format(Me.txt1(4).Text, "00") < Left(strSrvDate(1), 6) Then
            strSql = strSql & " And CP05<=" & Val(strMonthLastDate) & " "
            strSql = strSql & " And ((CP27 Is Null And CP57 Is Null) Or CP27>" & Val(strMonthLastDate) & " Or CP57>" & Val(strMonthLastDate) & " )"
            strSql = strSql & " And (EP14 Is Not Null And EP14<=" & Val(strMonthLastDate) & " ) "
            strSql = strSql & " And (EP15 Is Null Or EP15>" & Val(strMonthLastDate) & " ) "
        '若發文年月>=系統年月
        Else
            strSql = strSql & " AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL And EP29 Is Null "
        End If
    Else '個人
        strSql = strSql & " AND CP27 is null and cp57 is null and ep14 is not null AND EP15 IS NULL And EP29 Is Null"
    End If
    strSql = "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) select EP13,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) " & strSQL8 & strSql & " GROUP BY EP13 "
    cnnConnection.Execute strSql
    '可辦墨圖
    'Modify By Cheng 2003/07/07
'    strSQL = "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) select EP13,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P')  AND EP13='" & Trim(txt1(2).Text) & "' AND CP27 is null and cp57 is  null  and ep17 is not null and ep18 IS NULL GROUP BY EP13 "
'    strSQL = "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) select EP13,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 in ('CFP','P','FCP')  AND EP13='" & Trim(txt1(2).Text) & "' AND CP27 is null and cp57 is  null  and ep17 is not null and ep18 IS NULL And EP29 Is Null GROUP BY EP13 "
    strSql = ""
    If ProState = "2" Or ProState = "3" Then '管理
        '若發文年月<系統年月
        If Val(Me.txt1(3).Text) + 1911 & Format(Me.txt1(4).Text, "00") < Left(strSrvDate(1), 6) Then
            strSql = strSql & " And CP05<=" & Val(strMonthLastDate) & " "
            strSql = strSql & " And ((CP27 Is Null And CP57 Is Null) Or CP27>" & Val(strMonthLastDate) & " Or CP57>" & Val(strMonthLastDate) & " )"
            strSql = strSql & " And (EP17 Is Not Null And EP17<=" & Val(strMonthLastDate) & " ) "
            strSql = strSql & " And (EP18 Is Null Or EP18>" & Val(strMonthLastDate) & " ) "
        '若發文年月>=系統年月
        Else
            strSql = strSql & " AND CP27 is null and cp57 is  null  and ep17 is not null and ep18 IS NULL And EP29 Is Null "
        End If
    Else '個人
        strSql = strSql & " AND CP27 is null and cp57 is  null  and ep17 is not null and ep18 IS NULL And EP29 Is Null "
    End If
    'add by nick 2004/12/20   加多國案且草圖不計件不秀
    'edit by nickc 2005/03/01 墨圖也要判斷
    'StrSql = StrSql & " and ((cp21='Y' and ep20 is null) or cp21 is null) "
    strSql = strSql & " and ((cp21='Y' and (ep20 is null or ep29 is null)) or cp21 is null) "
      'add by nickc 2005/03/21
      strSql = strSql & " and  cp107='Y' "

'add by nickc 2005/04/18 以下作廢 因為改由管理那支統一算法
End With
Exit Sub
'
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) select EP13,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) " & strSQL8 & strSql & " GROUP BY EP13 "
'    cnnConnection.Execute strSql
'    '達成草圖
'    'Modify By Cheng 2003/07/07
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,count(*),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 IN ('P','CFP')  and cp57 is  null  and SUBSTR(EP15,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 GROUP BY EP13 "
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),SUM(eP16),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 IN ('P','CFP','FCP')  and cp57 is  null  and SUBSTR(EP15,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 GROUP BY EP13 "
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,ID) select EP13,3,Sum(Decode(EP29, Null, 1, 0)),sum(nvl(ep16,0)),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) and SUBSTR(EP15,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & " AND EP16 IS NOT NULL AND EP16>0 " & strSQL8 & " GROUP BY EP13 "
'    cnnConnection.Execute strSql
'    'Add By Cheng 2004/03/30
'    '將支援記錄計入達成草圖
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,ID) select SH02,3,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4)), 0,'" & strUserNum & "' from SupportHour, Staff S2 where SH02=S2.ST01(+) " & strSQL9 & " And SH11='V' GROUP BY SH02 "
'    cnnConnection.Execute strSql
'    'End
'    '達成墨圖
'    'Modify By Cheng 2003/07/07
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  cp01 IN ('P','CFP')  and cp57 is  null and SUBSTR(EP18,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & " GROUP BY EP13 "
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  cp01 IN ('P','CFP','FCP')  and cp57 is  null and SUBSTR(EP18,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & " GROUP BY EP13 "
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,4,Sum(Decode(EP29, Null, 1, 0)),sum(nvl(ep19,0)),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) and SUBSTR(EP18,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & strSQL8 & " GROUP BY EP13 "
'    cnnConnection.Execute strSql
'    'Add By Cheng 2004/03/30
'    '將支援記錄計入達成墨圖
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select SH02,4,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "' from SupportHour, Staff S2 where SH02=S2.ST01(+) " & strSQL9 & " And SH11='V' GROUP BY SH02 "
'    cnnConnection.Execute strSql
'    'End
'    '其他新案(抓CFP案草圖張數>0且草完日為當月的資料)
'    'Modify By Cheng 2003/07/07
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 NOT IN ('P','CFP')  AND SUBSTR(CP27,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND cp31='Y'  GROUP BY EP13 "
'    'Modify By Cheng 2004/02/13
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 NOT IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND cp31='Y'  GROUP BY EP13 "
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP20, Null, 1, 0)),SUM(EP16),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 IN ('CFP')  AND ((SUBSTR(CP27,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp57 is  null ) Or (SUBSTR(CP57,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and CP27 is  null) Or (CP27 Is Null Or CP57 Is Null)) And substr(EP15,1,6)=" & Val(Me.txt1(3).Text & Format(Me.txt1(4).Text, "00")) + 191100 & " And EP16>0  GROUP BY EP13 "
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,5,Sum(Decode(EP20, Null, 1, 0)),sum(nvl(ep16,0)),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress,Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND  CP01 IN ('CFP') And substr(EP15,1,6)=" & Val(Me.txt1(3).Text & Format(Me.txt1(4).Text, "00")) + 191100 & " And EP16>0 " & strSQL8 & " GROUP BY EP13 "
'    'End
'    cnnConnection.Execute strSql
'    '其他舊案(抓CFP案草圖張數<=0, 墨圖張數<=0且墨完日為當月的資料)
'    'Modify By Cheng 2003/07/07
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 NOT IN ('P','CFP')  AND SUBSTR(CP27,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND (CP31<>'Y' OR CP31 IS NULL) GROUP BY EP13 "
'    'Modify By Cheng 2004/02/13
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Txt1(2).Text) & "' AND  CP01 NOT IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Trim(Txt1(3).Text) & Format(Txt1(4).Text, "00")) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL AND (CP31<>'Y' OR CP31 IS NULL) GROUP BY EP13 "
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(EP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 IN ('CFP')  AND ((SUBSTR(CP27,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp57 is  null ) Or (SUBSTR(CP57,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp27 is  null) Or (CP27 Is Null Or CP57 Is Null)) And substr(EP18,1,6)=" & Val(Me.txt1(3).Text & Format(Me.txt1(4).Text, "00")) + 191100 & " and ((EP16<=0 Or EP16 Is Null) And (EP19<=0 Or EP19 Is Null)) GROUP BY EP13 "
'    'Modify by Morgan 2004/5/12
'    '不必判斷墨圖=0
'    'strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),SUM(EP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND  CP01 IN ('CFP') And substr(EP18,1,6)=" & Val(Me.txt1(3).Text & Format(Me.txt1(4).Text, "00")) + 191100 & " and ((EP16<=0 Or EP16 Is Null) And (EP19<=0 Or EP19 Is Null)) " & strSQL8 & " GROUP BY EP13 "
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,6,Sum(Decode(EP29, Null, 1, 0)),sum(nvl(ep19,0)),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND  CP01 IN ('CFP') And substr(EP18,1,6)=" & Val(Me.txt1(3).Text & Format(Me.txt1(4).Text, "00")) + 191100 & " and (EP16<=0 Or EP16 Is Null) " & strSQL8 & " GROUP BY EP13 "
'    'End
'    cnnConnection.Execute strSql
'    'Add By Cheng 2003/07/01
'    '本月發文
'    'Modify By Cheng 2003/07/07
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,count(*),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(txt1(2).Text) & "' AND  CP01 IN ('P','CFP')  AND SUBSTR(CP27,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL GROUP BY EP13 "
'    'Modify By Cheng 2003/07/17
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Txt1(2).Text) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Trim(Txt1(3).Text) & Format(Txt1(4).Text, "00")) + 191100 & "  and cp57 is  null and EP18 IS NOT NULL GROUP BY EP13 "
'    'Modify By Cheng 2004/02/16
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)),SUM(eP19),sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND EP13='" & Trim(Txt1(2).Text) & "' AND  CP01 IN ('P','CFP','FCP')  AND SUBSTR(CP27,1,6)=" & Val(Trim(Txt1(3).Text) & Format(Txt1(4).Text, "00")) + 191100 & "  and cp57 is  null GROUP BY EP13 "
'    '本月發文件數及點數
'    'Modify By Cheng 2004/03/30
'    '抓墨圖計件的資料
''    strSQL = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND SUBSTR(CP27,1,6)=" & Val(Trim(Txt1(3).Text) & Format(Txt1(4).Text, "00")) + 191100 & "  and cp57 is  null " & strSQL8 & " GROUP BY EP13 "
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9,Sum(Decode(EP29, Null, 1, 0)), 0, sum(cp18),'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND SUBSTR(CP27,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and cp57 is  null " & strSQL8 & " And EP29 Is Null GROUP BY EP13 "
'    'End
'    cnnConnection.Execute strSql
'    'Add By Cheng 2004/03/30
'    '將支援記錄計入本月發文件數
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select SH02,9,Sum(Decode(SH06, 'CFP', Nvl(SH05, 0)/8, Nvl(SH05, 0)/4)), 0, 0,'" & strUserNum & "' from SupportHour, Staff S2 where SH02=S2.ST01(+) " & strSQL9 & " And SH11='V' GROUP BY SH02 "
'    cnnConnection.Execute strSql
'    'End
'    '本月草圖張數/2(無墨完日或墨完日非當期)
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP16,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND SUBSTR(EP15,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & " and ( EP18 Is Null Or substr(EP18,1,6)<>" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & ") " & strSQL8 & " GROUP BY EP13 "
'    cnnConnection.Execute strSql
'    '本月墨圖張數/2(無草完日或草完日非當期)
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum(Nvl(EP19,0)/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND SUBSTR(EP18,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & " and ( EP15 Is Null Or substr(EP15,1,6)<>" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & ") " & strSQL8 & " GROUP BY EP13 "
'    cnnConnection.Execute strSql
'    '本月(草圖+墨圖張數)/2(草完日及墨完日皆為當期)
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,9, 0, Sum((Nvl(EP16,0)+Nvl(EP19,0))/2), 0,'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND SUBSTR(EP15,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  and substr(EP18,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & strSQL8 & " GROUP BY EP13 "
'    cnnConnection.Execute strSql
'    'End
'
'    'add by nickc 2005/04/13 增加提供圖檔及關聯
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,10,count(*), 0, 0,'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) And EP13=S2.ST01(+) AND SUBSTR(ep15,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  " & strSQL8 & " and ep20 is null GROUP BY EP13 "
'    cnnConnection.Execute strSql
'    strSql = "INSERT INTO R090706_1 (R111001,R111002,r111003,R111004,r111005,ID) select EP13,11,count(*), 0, 0,'" & strUserNum & "' from engineerprogress,caseprogress, Staff S2 where EP02=CP09(+) and ep20='N' and cp103=0.4 and cp100=0 And EP13=S2.ST01(+) AND SUBSTR(ep18,1,6)=" & Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100 & "  " & strSQL8 & " And EP29 Is Null GROUP BY EP13 "
'    cnnConnection.Execute strSql
'
'    DoEvents
'    Select Case ProState
'    Case "2", "3"
'        strBeginDate = ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString((Val(Me.txt1(3).Text) + 1911) & Format(Me.txt1(4).Text, "00") & "01")))
'        strEndDate = IIf((Val(Me.txt1(3).Text) + 1911) & Format(Me.txt1(4).Text, "00") < Left(strSrvDate(1), 6), (Val(Me.txt1(3).Text) + 1911) & Format(Me.txt1(4).Text, "00") & PUB_GetMonthDays(Val(Me.txt1(3).Text) + 1911, Val(Me.txt1(4).Text)), strSrvDate(1))
'    Case Else
'        strBeginDate = ChangeWDateStringToWString(DateAdd("m", -1, ChangeWStringToWDateString(Left(strSrvDate(1), 6) & "01")))
'        strEndDate = strSrvDate(1)
'    End Select
'    Set NickRS = New ADODB.Recordset
'    'Modify By Cheng 2003/06/30
''    strSQL = "SELECT EP13,PA08,EP14,'1' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Txt1(2).Text) & "' AND CP57 IS NULL AND CP27 is null and ep14 is not null AND EP15 IS NULL "
''    strSQL = strSQL & " UNION all  SELECT EP13,PA08,EP17,'2' FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT WHERE CP01 IN ('P','CFP') AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(Txt1(2).Text) & "' AND CP57 IS NULL AND CP27 is null and ep17 is not null and ep18 IS NULL "
''    strSQL = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2  WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(txt1(2).Text) & "' AND CP57 IS NULL and ep14 is not null And EP20 Is Null " & StrSQL6
''    strSQL = strSQL & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2 WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(txt1(2).Text) & "' AND CP57 IS NULL and (ep17 is not null Or EP08 Is Not Null) And EP29 Is Null " & StrSQL6
''edit by nick 2005/01/12 morgan 之前忘記改
''    StrSql = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2  WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) And EP20 Is Null And (EP14>=" & strBeginDate & " And EP14<=" & strEndDate & " ) " & StrSQL7
''    StrSql = StrSql & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2 WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) And EP29 Is Null And (EP17>=" & strBeginDate & " And EP17<=" & strEndDate & ") " & StrSQL7
'    strSql = "SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57,pa08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2  WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) And EP20 Is Null And (EP14>=" & strBeginDate & " And EP14<=" & strEndDate & " ) " & StrSQL7
'    strSql = strSql & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57,pa08 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2 WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) And EP29 Is Null And (EP17>=" & strBeginDate & " And EP17<=" & strEndDate & ") " & StrSQL7
'
''    If ProState <> 4 Then
''        strSQL = strSQL & " Union SELECT EP13, CP10, EP14, '1', EP15, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2 WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(txt1(2).Text) & "' AND CP57 IS NULL and ep14 is not null And EP20 Is Null " & strSQL1
''        strSQL = strSQL & " Union SELECT EP13, CP10, Nvl(EP17, EP08), '2', EP18, CP07, CP27, CP09, CP57 FROM ENGINEERPROGRESS,CASEPROGRESS,PATENT,Staff S2 WHERE EP13=S2.ST01(+) AND EP02=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND EP13='" & Trim(txt1(2).Text) & "' AND CP57 IS NULL and (ep17 is not null Or EP08 Is Not Null) And EP29 Is Null " & strSQL1
''    End If
'    NickRS.CursorLocation = adUseClient
'    NickRS.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If NickRS.RecordCount <> 0 Then
'        NickRS.MoveFirst
'        Do While NickRS.EOF = False
''edit by nick 2005/01/12
''            Select Case CheckStr(NickRS.Fields(1))
''                Case "103", "105" '設計申請
'            Select Case CheckStr(NickRS.Fields("PA08").Value)
'               Case "3"
'
'                    If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
'                        '若有草齊日及草完日
'                        If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
'                            '草完日必須為當月
'                            'edit by nick 2005/01/12 程弘忘記改
'                            'If Left("" & NickRS.Fields(4).Value, 6) = "" & (Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100) Then
'                            If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.txt1(4).Text) + 191100), Left(strSrvDate(1), 6)) Then
'                                If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 5 Then
'                                    cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
'                                End If
'                            End If
''                        '若無發文日有草齊日無草完日無消收文日
''                        ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
'                        '若有草齊日無草完日
'                        ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
''                            If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 5 Then
'                            If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 5 Then
'                                cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
'                            End If
'                        End If
'                    Else '墨圖
'                        '若有墨齊日及墨完日
'                        If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
'                            '墨完日必須為當月
'                            'edit by nick 2005/01/12 程弘忘記改
'                            'If Left("" & NickRS.Fields(4).Value, 6) = "" & (Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100) Then
'                            If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.txt1(4).Text) + 191100), Left(strSrvDate(1), 6)) Then
'                                If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 3 Then
'                                    cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
'                                End If
'                            End If
''                        '若無發文日有墨齊日無墨完日無取消收文日
''                        ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
'                        '若有墨齊日無墨完日
'                        ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
''                            If GetWorkDay(strSrvDate(1), CheckStr(NickRS.Fields(2))) > 3 Then
'                            If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 3 Then
'                                cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
'                            End If
'                        End If
'                    End If
'                Case Else
'                    If CheckStr(NickRS.Fields(3)) = "1" Then  '草圖
'                        '若有草齊日及草完日
'                        If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
'                            '草完日必須為當月
'                            'edit by nick 2005/01/12 程弘忘記改
'                            'If Left("" & NickRS.Fields(4).Value, 6) = "" & (Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100) Then
'                            If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.txt1(4).Text) + 191100), Left(strSrvDate(1), 6)) Then
'                                If GetWorkDay(CheckStr(NickRS.Fields(4)), "" & NickRS.Fields(2).Value) > 4 Then
'                                    cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
'                                End If
'                            End If
''                        '若無發文日有草齊日無草完日無取消收文日
''                        ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
'                        '若有草齊日無草完日
'                        ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
'                            If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 4 Then
'                                cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',7,1,'" & strUserNum & "') "
'                            End If
'                        End If
'                    Else '墨圖
'                        '若有墨齊日及墨完日
'                        If "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value <> "" Then
'                            '墨完日必須為當月
'                            'edit by nick 2005/01/12 程弘忘記改
'                            'If Left("" & NickRS.Fields(4).Value, 6) = "" & (Val(Trim(txt1(3).Text) & Format(txt1(4).Text, "00")) + 191100) Then
'                            If Left("" & NickRS.Fields(4).Value, 6) = IIf(ProState = "2" Or ProState = "3", "" & (Val(Me.txt1(4).Text) + 191100), Left(strSrvDate(1), 6)) Then
'                                If GetWorkDay((CheckStr(NickRS.Fields(4))), "" & NickRS.Fields(2).Value) > 3 Then
'                                    cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
'                                End If
'                            End If
''                        '若無發文日有墨齊日無墨完日無取消收文日
''                        ElseIf "" & NickRS.Fields(6).Value = "" And "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" And "" & NickRS.Fields(8).Value = "" Then
'                        '若有墨齊日無墨完日
'                        ElseIf "" & NickRS.Fields(2).Value <> "" And "" & NickRS.Fields(4).Value = "" Then
'                            If GetWorkDay(strEndDate, CheckStr(NickRS.Fields(2))) > 3 Then
'                                cnnConnection.Execute "INSERT INTO R090706_1 (R111001,R111002,R111003,ID) VALUES ('" & CheckStr(NickRS.Fields(0)) & "',8,1,'" & strUserNum & "') "
'                            End If
'                        End If
'                    End If
'            End Select
'           NickRS.MoveNext
'        Loop
'    End If
'    If NickRS.State = 1 Then NickRS.Close
'
'    '寫入沒有資料的繪圖人員
'    CheckOC
'    strSql = "SELECT DISTINCT R111001 FROM R090706_1 WHERE ID='" & strUserNum & "' "
'    .CursorLocation = adUseClient
'    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 And .RecordCount > 0 Then
'        .MoveFirst
'        Do While .EOF = False
'            For i = 1 To 8
'                strSql = "SELECT * FROM R090706_1 WHERE ID='" & strUserNum & "' AND R111002='" & Trim(str(i)) & "' AND R111001='" & CheckStr(.Fields(0)) & "' "
'                CheckOC2
'                adoRecordset1.CursorLocation = adUseClient
'                adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'                If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
'                Else
'                    strSql = "INSERT INTO R090706_1 VALUES('" & CheckStr(.Fields(0)) & "','" & Trim(str(i)) & "',0,0,0,'" & strUserNum & "') "
'                    cnnConnection.Execute strSql
'                End If
'            Next i
'            .MoveNext
'            CheckOC2
'        Loop
'    End If
'   CheckOC
'End With
End Sub

Sub Process1()
If Val(txt1(5)) = 1 Then
    Me.Hide
    frm090706_1.Show
Else
    PrintData
End If
End Sub

Sub PrintData()
'strSQL = "SELECT DISTINCT R110001 FROM R090706 WHERE ID='" & strUserNum & "' "
strSql = "SELECT R110001,ST02 FROM R090706,STAFF WHERE R110001=ST01(+) AND ID='" & strUserNum & "' GROUP BY R110001,ST02"
CheckOC2
Page = 1
adoRecordset1.CursorLocation = adUseClient
adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
    adoRecordset1.MoveFirst
    Do While adoRecordset1.EOF = False
'        strTemp3 = CheckStr(adoRecordset1.Fields(0))
        strTemp3 = CheckStr(adoRecordset1.Fields(1))
        'add by nickc 2005/04/23 因為 frm090711 已經抓過資料了，所以直接使用，不在抓了
        PrintData1 (CheckStr(adoRecordset1.Fields(0)))
        'PrintData2 (CheckStr(adoRecordset1.Fields(0)))
        PrintEnd1 (CheckStr(adoRecordset1.Fields(0)))
        Page = Page + 1
        Printer.NewPage
        adoRecordset1.MoveNext
    Loop
End If
CheckOC2
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintData1(Strindex As String)
If Len(Strindex) = 0 Then
   'Modify By Cheng 2002/04/19
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') "
    'Modify By Cheng 2003/06/30
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') Order By To_Number(Replace(R110003,'/','')) "
    'edit by nickc 2005/05/12
    'strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') Order By R110026 Desc , R110006 Desc "
    strSql = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') Order By decode(R110003,null,'     ' ,r110003) Desc,R110006"
Else
   'Modify By Cheng 2002/04/19
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' "
    'Modify By Cheng 2003/06/30
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' Order By To_Number(Replace(R110003,'/','')) "
    'edit by nickc 2005/05/12
    'strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' Order By R110026 Desc, R110006 Desc "
    strSql = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' Order By decode(R110003,null,'     ',r110003) Desc,R110006"
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 22
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Add By Cheng 2002/04/17
            strTemp(23) = CheckStr(.Fields(24))
            strTemp(3) = StrToStr(strTemp(3), 3)
            strTemp(6) = StrToStr(strTemp(6), 7)
            strTemp(7) = StrToStr(strTemp(7), 4)
            'Modify By Cheng 2002/04/17
'            strTemp(22) = StrToStr(strTemp(20), 8)
            strTemp(22) = StrToStr(strTemp(22), 8)
            'Add By Cheng 2003/06/30
            strTemp(24) = CheckStr(.Fields(25))
            PrintDatil
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintEnd1(Strindex As String)
'列印結尾
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
End If

If Len(Strindex) = 0 Then
   'Modify By Cheng 2002/04/17
'    strSQL = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)) FROM R090706_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') group by r111002 order by r111002 "
    'edit by nick 2005/01/12
    'StrSql = "SELECT r111002,SUM(NVL(R111003,0)),SUM(NVL(R111004,0)),SUM(NVL(R111005,0)) FROM R090706_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') group by r111002 order by r111002 "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT r111002,SUM(NVL(R111003,0)),SUM(NVL(R111004,0)),SUM(NVL(R111005,0)) FROM R090711_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') group by r111002 order by r111002 "
    strSql = "SELECT r111002,SUM(NVL(R111003,0)),SUM(NVL(R111004,0)),SUM(NVL(R111005,0)),SUM(NVL(R111006,0)),SUM(NVL(R111007,0)),SUM(NVL(R111008,0)) FROM R090711_1 WHERE ID='" & strUserNum & "' AND (R111001 IS NULL OR R111001='') group by r111002 order by r111002 "
Else
   'Modify By Cheng 2002/04/17
'    strSQL = "SELECT r111002,SUM(DECODE(R111003,0,0,R111003)),SUM(DECODE(R111004,0,0,R111004)),SUM(DECODE(R111005,0,0,R111005)) FROM R090706_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' group by r111002 order by r111002 "
    'edit by nick 2005/01/12
    'StrSql = "SELECT r111002,SUM(NVL(R111003,0)),SUM(NVL(R111004,0)),SUM(NVL(R111005,0)) FROM R090706_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' group by r111002 order by r111002 "
    'edit by nickc 2005/05/04
    'strSQL = "SELECT r111002,SUM(NVL(R111003,0)),SUM(NVL(R111004,0)),SUM(NVL(R111005,0)) FROM R090711_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' group by r111002 order by r111002 "
    strSql = "SELECT r111002,SUM(NVL(R111003,0)),SUM(NVL(R111004,0)),SUM(NVL(R111005,0)),SUM(NVL(R111006,0)),SUM(NVL(R111007,0)),SUM(NVL(R111008,0)) FROM R090711_1 WHERE ID='" & strUserNum & "' AND R111001='" & Strindex & "' group by r111002 order by r111002 "
End If
CheckOC
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "當月統計："
iPrint = iPrint + 300
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        Do While .EOF = False
            Select Case Val(CheckStr(.Fields(0)))
            Case 1
                 Printer.CurrentX = 0
                 Printer.CurrentY = iPrint
                 'edit by nickc 2005/05/04
                 'Printer.Print "可辦草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件"
                 Printer.Print "可辦草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,###,##0.00") & ") 件"
            Case 2
                 Printer.CurrentX = 4000
                 Printer.CurrentY = iPrint
                 'edit by nickc 2005/05/04
                 'Printer.Print "可辦墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件"
                 Printer.Print "可辦墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,###,##0.00") & ") 件"
            Case 3
                 Printer.CurrentX = 8000
                 Printer.CurrentY = iPrint
                 'edit by nickc 2005/05/04
                 'Printer.Print "達成草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0.0") & " 張"
                 Printer.Print "達成草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,###,##0.00") & ") 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0.0") & " 張"
            Case 4
                 Printer.CurrentX = 12000
                 Printer.CurrentY = iPrint
                 'edit by nickc 2005/05/04
                 'Printer.Print "達成墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0.0") & " 張 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & " 點 "
                 Printer.Print "達成墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,###,##0.00") & ") 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0.0") & " 張 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(6)), "###,###,###,###,##0.00") & ") 點 "
            Case 5
                 Printer.CurrentX = 0
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "其他新案:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & " 點 "
            Case 6
                 Printer.CurrentX = 4000
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "其他舊案:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & " 點"
            Case 7
                 Printer.CurrentX = 8000
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "逾時草圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件"
            Case 8
                 Printer.CurrentX = 12000
                 Printer.CurrentY = iPrint + 300
                 Printer.Print "逾時墨圖:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件"
            'Add By Cheng 2003/07/01
            Case 9
                 Printer.CurrentX = 0
                 Printer.CurrentY = iPrint + 600
                 'edit by nickc 2005/05/04
                 'Printer.Print "本月發文:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0.0") & " 張 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & " 點 "
                 Printer.Print "本月發文:" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(4)), "###,###,###,###,##0.00") & ") 件 " & Format(CheckStr(.Fields(2)), "###,###,###,###,##0.0") & " 張 " & Format(CheckStr(.Fields(3)), "###,###,###,###,##0.00") & "(" & Format(CheckStr(.Fields(6)), "###,###,###,###,##0.00") & ") 點 "
            'add by nickc 2005/04/13 加入提供圖黨及轉換 0.4
            Case 10
                 Printer.CurrentX = 8000
                 Printer.CurrentY = iPrint + 600
                 Printer.Print "提供圖檔(0.6):" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件 "
            Case 11
                 Printer.CurrentX = 12000
                 Printer.CurrentY = iPrint + 600
                 Printer.Print "轉換案(0.4):" & Format(CheckStr(.Fields(1)), "###,###,###,###,##0.00") & " 件 "
                 
            Case Else
            End Select
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub

Sub PrintTitle() '列印抬頭

iPrint = 0
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 5500
Printer.CurrentY = iPrint
Printer.Print "繪圖人員工作進度資料表"
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
iPrint = iPrint + 500
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
'Modify By Cheng 2003/05/28
'Printer.Print "年月：" & Txt1(3) & "/" & Txt1(4)
Printer.Print "年月：" & Val(Format(txt1(3), "000")) & "/" & Format(txt1(4), "00")
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
'Modify By Cheng 2003/05/28
'Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
Printer.Print "列印日期：" & Format(strSrvDate(2), "###/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 0
Printer.CurrentY = iPrint
Printer.Print "繪圖人員：" & strTemp3
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁　　次：" & str(Page)
iPrint = iPrint + 300
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
GetPleft

Printer.Font.Size = 9
Printer.CurrentX = PLeft(1)
'edit by nickc 2005/05/12 類別不秀
'Printer.CurrentY = iPrint
''Printer.Print "收文"
'Printer.Print "類"
'Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
'edit by nickc 2005/05/12 改秀發文日
'Printer.Print "收文日"
Printer.Print "發文日"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "種類"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "點數"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
'Add By Cheng 2002/04/17
Printer.CurrentX = PLeft(23)
Printer.CurrentY = iPrint
Printer.Print "草"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "承辦人"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "草　圖"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "草　圖"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "草圖"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "草圖作"
'Add By Cheng 2003/06/30
Printer.CurrentX = PLeft(24)
Printer.CurrentY = iPrint
Printer.Print "墨"
Printer.CurrentX = PLeft(12) + 100
Printer.CurrentY = iPrint
Printer.Print "墨　圖"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "墨　圖"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "墨圖"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "墨圖作"
Printer.CurrentX = PLeft(16) + 100
Printer.CurrentY = iPrint
Printer.Print "承辦 時數"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "複雜"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
'Printer.Print "修  改  時  數"
Printer.Print "修　改　時　數"
Printer.CurrentX = PLeft(22)
Printer.CurrentY = iPrint
Printer.Print "備註"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.CurrentX = PLeft(1)
'edit by nickc 2005/05/12 類別不秀
'Printer.CurrentY = iPrint
''Printer.Print "類別"
'Printer.Print "別"
'Add By Cheng 2002/04/17
Printer.CurrentX = PLeft(23)
Printer.CurrentY = iPrint
Printer.Print "計"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "張數"
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "業天數"
Printer.CurrentX = PLeft(24)
Printer.CurrentY = iPrint
Printer.Print "計"
Printer.CurrentX = PLeft(12) + 100
Printer.CurrentY = iPrint
Printer.Print "齊備日"
Printer.CurrentX = PLeft(13)
Printer.CurrentY = iPrint
Printer.Print "完稿日"
Printer.CurrentX = PLeft(14)
Printer.CurrentY = iPrint
Printer.Print "張數"
Printer.CurrentX = PLeft(15)
Printer.CurrentY = iPrint
Printer.Print "業天數"
Printer.CurrentX = PLeft(16) + 100
Printer.CurrentY = iPrint
Printer.Print "草圖"
Printer.CurrentX = PLeft(17)
Printer.CurrentY = iPrint
Printer.Print "墨圖"
Printer.CurrentX = PLeft(18)
Printer.CurrentY = iPrint
Printer.Print "件數"
Printer.CurrentX = PLeft(19)
Printer.CurrentY = iPrint
Printer.Print "1"
Printer.CurrentX = PLeft(20)
Printer.CurrentY = iPrint
Printer.Print "2"
Printer.CurrentX = PLeft(21)
Printer.CurrentY = iPrint
Printer.Print "3"
iPrint = iPrint + 300
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
ShowLine
If iPrint >= 9000 Then
    Page = Page + 1
    Printer.NewPage
    PrintTitle
    Exit Sub
End If
Printer.Font.Size = 9
End Sub

Sub PrintDatil() '列印資料
'For i = 1 To 22
'For i = 1 To 23
For i = 1 To 24
    Select Case i
    'add by nickc 2005/05/12
    Case 1   '類別不秀
        
    Case 4
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0.00"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "##0.00")
'    Case 10, 11, 14, 15, 16, 17, 18, 19, 20, 21
'    Case 10, 11, 13, 14, 15, 16, 17, 18, 19, 20, 21
'    Case 10, 11, 14, 15, 16, 17, 18, 19, 20, 21
    Case 10, 11, 14, 15, 18
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "##0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "##0")
    Case 16, 17, 19, 20, 21
        Printer.CurrentX = PLeft(i) + 300 - Printer.TextWidth(Format(strTemp(i), "#.0"))
        Printer.CurrentY = iPrint
        Printer.Print Format(strTemp(i), "#.0")
    Case Else
        Printer.CurrentX = PLeft(i)
        Printer.CurrentY = iPrint
        Printer.Print strTemp(i)
    End Select
Next i
iPrint = iPrint + 300
End Sub

Sub GetPleft()
'定陣列
'字 SIZE = 9
'1 WORD = 180 PIX
'0.5 WORD = 90 PIX
'SPACE = 90 PIX
Erase PLeft
PLeft(0) = 0
PLeft(1) = 0 '收文類別
PLeft(2) = PLeft(1) + (1.5 * 180) '收文日
PLeft(3) = PLeft(2) + (4.5 * 180) '種類
PLeft(4) = PLeft(3) + (5.5 * 180) - 200 '點數
PLeft(5) = PLeft(4) + (2.5 * 180) + 50 '本所案號
PLeft(6) = PLeft(5) + (8 * 180) '案件名稱
PLeft(7) = PLeft(6) + (8 * 180) '承辦人
PLeft(23) = PLeft(7) + (4.5 * 180) '草圖是否計件
'PLeft(8) = PLeft(7) + (4.5 * 180) '草圖齊備日
PLeft(8) = PLeft(23) + (1.5 * 180) '草圖齊備日
PLeft(9) = PLeft(8) + (4.5 * 180) '草圖完稿日
PLeft(10) = PLeft(9) + (4 * 180) '草圖張數
PLeft(11) = PLeft(10) + (3 * 180) '草圖作天數
PLeft(24) = PLeft(11) + (3.5 * 180) '墨圖是否計件
'PLeft(12) = PLeft(11) + (3 * 180) '墨圖齊備日
PLeft(12) = PLeft(24) + (1.5 * 180) '墨圖齊備日
PLeft(13) = PLeft(12) + (4.5 * 180) '墨圖完稿日
PLeft(14) = PLeft(13) + (4.5 * 180) '墨圖張數
PLeft(15) = PLeft(14) + (3 * 180) '墨圖作業天數
PLeft(16) = PLeft(15) + (3 * 180) '承辦時數(草圖)
PLeft(17) = PLeft(16) + (3 * 180) '承辦時數(墨圖)
PLeft(18) = PLeft(17) + (3 * 180) '複雜件數
PLeft(19) = PLeft(18) + (3 * 180) '修改時數1
PLeft(20) = PLeft(19) + (3 * 180) '修改時數2
PLeft(21) = PLeft(20) + (3 * 180) '修改時數3
PLeft(22) = PLeft(21) + (3 * 180) '備註
End Sub


Sub ShowLine()
Printer.Line (0, iPrint + 150)-(16500, iPrint + 150)
iPrint = iPrint + 300
End Sub

Private Sub Form_Activate()
ProState = m_ProState 'Add By Sindy 2017/8/10 重新設定權限
End Sub

Private Sub Form_Load()
m_ProState = ProState 'Add By Sindy 2017/8/10 記錄目前權限
MoveFormToCenter Me

Select Case ProState
Case "1"  '個人
   txt1(5) = "2"
   txt1(2) = strUserNum
   txt1(3) = Val(Mid(strSrvDate(1), 1, 4)) - 1911
   txt1(4) = Mid(strSrvDate(1), 5, 2)
   txt1(6) = ""
   Frame1.Visible = False
   Me.Caption = "繪圖人員工作進度資料列印"
   Me.Height = 1355
    SendKeys "{Tab}"
Case "2"  '管理
    'Add By Cheng 2004/02/18
    '發文年月預設為系統年月
    txt1(3) = Val(Mid(strSrvDate(1), 1, 4)) - 1911
    txt1(4) = Mid(strSrvDate(1), 5, 2)
    'End
   txt1(5) = "1"
   txt1(6) = "Y"
   Me.Caption = "繪圖人員工作進度資料查詢"
   Frame1.Visible = True
   Me.Height = 2955
    SendKeys "{Tab}"
Case Else
End Select

lbl1(0).Caption = ""  'Added by Lydia 2022/01/28
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm090706 = Nothing
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
Case 0
     Select Case Trim(txt1(0))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(0).SetFocus
          txt1(0).SelStart = 0
          txt1(0).SelLength = Len(txt1(0))
          Exit Sub
     End Select
Case 1
     Select Case Trim(txt1(1))
     Case "1", "2", "", "3", "4", "5"
     Case Else
          s = MsgBox("所別只能輸入 1 到 5 !!", , "USER 輸入錯誤")
          txt1(1).SetFocus
          txt1(1).SelStart = 0
          txt1(1).SelLength = Len(txt1(1))
          Exit Sub
     End Select
Case 2
     lbl1(0).Caption = GetPrjSales(txt1(2))
Case 3
     If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
            s = MsgBox("年輸入錯誤!!", , "USER 輸入錯誤")
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
     End If
Case 4
        If IsNumeric(txt1(Index)) = False And Len(txt1(Index)) <> 0 Then
            s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
            txt1(Index).SetFocus
            txt1(Index).SelStart = 0
            txt1(Index).SelLength = Len(txt1(Index))
            Exit Sub
        Else
            If Val(txt1(Index)) < 1 Or Val(txt1(Index)) > 12 Then
                s = MsgBox("月輸入錯誤!!", , "USER 輸入錯誤")
                txt1(Index).SetFocus
                txt1(Index).SelStart = 0
                txt1(Index).SelLength = Len(txt1(Index))
                Exit Sub
            End If
         End If
Case 5
     Select Case Trim(txt1(5))
     Case "1", "2", ""
     Case Else
          s = MsgBox("顯示方式只能輸入 1 或 2 !!", , "USER 輸入錯誤")
          txt1(5).SetFocus
          txt1(5).SelStart = 0
          txt1(5).SelLength = Len(txt1(5))
          Exit Sub
     End Select
Case 6
     Select Case Trim(txt1(6))
     Case "Y", "N", ""
     Case Else
          s = MsgBox("只能輸入 Y 或 N 或 空白！", , "USER 輸入錯誤")
          txt1(6).SetFocus
          txt1(6).SelStart = 0
          txt1(6).SelLength = Len(txt1(6))
          Exit Sub
     End Select
Case Else
End Select
End Sub

'add by nickc 2005/04/23 因為 frm090711 已經抓過資料了，所以直接使用，不在抓了
Sub PrintData2(Strindex As String)
If Len(Strindex) = 0 Then
   'Modify By Cheng 2002/04/19
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') "
    'Modify By Cheng 2003/06/30
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') Order By To_Number(Replace(R110003,'/','')) "
    strSql = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND (R110001 IS NULL OR R110001='') Order By R110026 Desc , R110006 Desc "
Else
   'Modify By Cheng 2002/04/19
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' "
    'Modify By Cheng 2003/06/30
'    strSQL = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' Order By To_Number(Replace(R110003,'/','')) "
    strSql = "SELECT * FROM R090706_2 WHERE ID='" & strUserNum & "' AND R110001='" & Strindex & "' Order By R110026 Desc, R110006 Desc "
End If
CheckOC
With adoRecordset
    .CursorLocation = adUseClient
    .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If .RecordCount <> 0 And .RecordCount > 0 Then
        .MoveFirst
        PrintTitle
        Do While .EOF = False
            For i = 0 To 22
                strTemp(i) = CheckStr(.Fields(i))
            Next i
            'Add By Cheng 2002/04/17
            strTemp(23) = CheckStr(.Fields(24))
            strTemp(3) = StrToStr(strTemp(3), 3)
            strTemp(6) = StrToStr(strTemp(6), 7)
            strTemp(7) = StrToStr(strTemp(7), 4)
            'Modify By Cheng 2002/04/17
'            strTemp(22) = StrToStr(strTemp(20), 8)
            strTemp(22) = StrToStr(strTemp(22), 8)
            'Add By Cheng 2003/06/30
            strTemp(24) = CheckStr(.Fields(25))
            PrintDatil
            If iPrint >= 9000 Then
                Page = Page + 1
                Printer.NewPage
                PrintTitle
            End If
            .MoveNext
        Loop
    End If
End With
CheckOC
End Sub
