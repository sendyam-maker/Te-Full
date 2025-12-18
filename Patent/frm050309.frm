VERSION 5.00
Begin VB.Form frm050309 
   BorderStyle     =   1  '單線固定
   Caption         =   "承辦人發文明細表"
   ClientHeight    =   3300
   ClientLeft      =   3120
   ClientTop       =   1530
   ClientWidth     =   3135
   ControlBox      =   0   'False
   LinkTopic       =   "Form8"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   3135
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   12
      Left            =   2244
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2610
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   11
      Left            =   924
      MaxLength       =   9
      TabIndex        =   10
      Top             =   2610
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   10
      Left            =   1764
      MaxLength       =   1
      TabIndex        =   9
      Top             =   2310
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   14
      Left            =   2244
      MaxLength       =   9
      TabIndex        =   13
      Top             =   2910
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   13
      Left            =   924
      MaxLength       =   9
      TabIndex        =   12
      Top             =   2910
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   8
      Left            =   924
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1980
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   7
      Left            =   924
      MaxLength       =   1
      TabIndex        =   7
      Top             =   1680
      Width           =   255
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   6
      Left            =   2244
      MaxLength       =   4
      TabIndex        =   6
      Top             =   1380
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   5
      Left            =   924
      MaxLength       =   4
      TabIndex        =   5
      Top             =   1380
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   4
      Left            =   2244
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   3
      Left            =   924
      MaxLength       =   4
      TabIndex        =   3
      Top             =   1080
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   2
      Left            =   2244
      MaxLength       =   7
      TabIndex        =   2
      Top             =   780
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   1
      Left            =   924
      MaxLength       =   7
      TabIndex        =   1
      Top             =   780
      Width           =   800
   End
   Begin VB.TextBox txt1 
      Height          =   264
      Index           =   0
      Left            =   924
      TabIndex        =   0
      Top             =   480
      Width           =   2136
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   400
      Index           =   1
      Left            =   2205
      TabIndex        =   15
      Top             =   24
      Width           =   800
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   1410
      TabIndex        =   14
      Top             =   24
      Width           =   756
   End
   Begin VB.Line Line5 
      X1              =   1890
      X2              =   2130
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line4 
      X1              =   1890
      X2              =   2130
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line3 
      X1              =   1884
      X2              =   2124
      Y1              =   1500
      Y2              =   1500
   End
   Begin VB.Line Line2 
      X1              =   1884
      X2              =   2124
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Line Line1 
      X1              =   1884
      X2              =   2124
      Y1              =   900
      Y2              =   900
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "申請人："
      Height          =   180
      Left            =   90
      TabIndex        =   27
      Top             =   2610
      Width           =   720
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "(Y 列印)"
      Height          =   180
      Left            =   2130
      TabIndex        =   26
      Top             =   2310
      Width           =   645
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "是否列印明細："
      Height          =   180
      Left            =   90
      TabIndex        =   25
      Top             =   2310
      Width           =   1260
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "代理人："
      Height          =   180
      Left            =   90
      TabIndex        =   24
      Top             =   2910
      Width           =   720
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Height          =   180
      Left            =   1776
      TabIndex        =   23
      Top             =   1980
      Width           =   1080
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "承辦人："
      Height          =   180
      Left            =   90
      TabIndex        =   22
      Top             =   1980
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(1. 承辦人 2. 代理人)"
      Height          =   180
      Left            =   1404
      TabIndex        =   21
      Top             =   1680
      Width           =   1608
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "列印順序："
      Height          =   180
      Left            =   90
      TabIndex        =   20
      Top             =   1680
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "案件性質："
      Height          =   180
      Left            =   90
      TabIndex        =   19
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "申請國家："
      Height          =   180
      Left            =   90
      TabIndex        =   18
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "發文日："
      Height          =   180
      Left            =   90
      TabIndex        =   17
      Top             =   780
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "系統類別："
      Height          =   180
      Left            =   90
      TabIndex        =   16
      Top             =   480
      Width           =   900
   End
End
Attribute VB_Name = "frm050309"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/6 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/12 日期欄已修改
Option Explicit
Dim strSql As String, i As Integer, j As Integer, s As Integer, k As Integer
Dim strTemp1 As Variant, strTemp2 As Variant, StrTest As String, StrTest2 As String
Dim strSQL1 As String, strSQL2 As String, strTemp(0 To 11) As String
Dim PLeft(0 To 11) As Integer, Page As Integer, iPrint As Integer
Dim StrTemp6 As String, StrTemp4(0 To 4) As String, StrTemp5(0 To 4) As String
'Add By Cheng 2002/09/16
Dim blnClkSure As Boolean '判斷是否按下確定按鈕
'Add By Cheng 2003/04/14
Dim m_dblSubTotal As Double '件數合計
 
Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0 '確定
   'Add By Cheng 2002/09/16
   blnClkSure = False
     If Len(txt1(0)) = 0 Then
        s = MsgBox("系統類別不可空白!!", , "USER 輸入錯誤")
        txt1(0).SetFocus
        txt1(0).SelStart = 0
        txt1(0).SelLength = Len(txt1(0))
        Exit Sub
     Else
        If Len(txt1(2)) = 0 Then
            s = MsgBox("發文日不可空白!!", , "USER 輸入錯誤")
            txt1(1).SetFocus: txt1(1).SelStart = 0: txt1(1).SelLength = Len(txt1(1))
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
                  MsgBox "發文日範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(1).SetFocus
                  txt1_GotFocus 1
                  Exit Sub
               End If
            End If
            If Me.txt1(3).Text <> "" And Me.txt1(4).Text <> "" Then
               If Me.txt1(3).Text > Me.txt1(4).Text Then
                  MsgBox "申請國家範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(3).SetFocus
                  txt1_GotFocus 3
                  Exit Sub
               End If
            End If
            If Me.txt1(5).Text <> "" And Me.txt1(6).Text <> "" Then
               If Me.txt1(5).Text > Me.txt1(5).Text Then
                  MsgBox "案件性質範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                  blnClkSure = True
                  Me.txt1(5).SetFocus
                  txt1_GotFocus 5
                  Exit Sub
               End If
            End If
            
            If Len(txt1(7)) = 0 Then
                s = MsgBox("列印順序不可空白!!", , "USER 輸入錯誤")
                txt1(7).SetFocus
                txt1(7).SelStart = 0
                txt1(7).SelLength = Len(txt1(7))
                Exit Sub
            Else
               'Add By Cheng 2002/09/16
               lbl1 = GetPrjSales(txt1(8))
               If Me.txt1(8).Text <> "" Then
                  If Me.txt1(8).Text = Me.lbl1.Caption Then
                     Me.lbl1.Caption = ""
                     Me.txt1(8).SetFocus
                     txt1_GotFocus 8
                     Exit Sub
                  End If
               End If
               If Len(txt1(11)) <> 0 Then
                  If Left(txt1(11), 6) <> Left(txt1(12), 6) Then
                      s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
                      blnClkSure = True
                      txt1(11).SetFocus
                      txt1_GotFocus 11
                      Exit Sub
                  End If
               End If
               If Me.txt1(11).Text <> "" And Me.txt1(12).Text <> "" Then
                  If Me.txt1(11).Text > Me.txt1(12).Text Then
                     MsgBox "申請人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(11).SetFocus
                     txt1_GotFocus 11
                     Exit Sub
                  End If
               End If
               If Len(txt1(13)) <> 0 Then
                  If Left(txt1(13), 6) <> Left(txt1(14), 6) Then
                      s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
                      blnClkSure = True
                      txt1(13).SetFocus
                      txt1_GotFocus 13
                      Exit Sub
                  End If
               End If
               If Me.txt1(13).Text <> "" And Me.txt1(14).Text <> "" Then
                  If Me.txt1(13).Text > Me.txt1(14).Text Then
                     MsgBox "代理人範圍輸入錯誤!!!", vbExclamation + vbOKOnly
                     blnClkSure = True
                     Me.txt1(13).SetFocus
                     txt1_GotFocus 13
                     Exit Sub
                  End If
               End If
            
                Me.Enabled = False
                Process
                Me.Enabled = True
            End If
        End If
     End If
Case 1
     Unload Me
Case Else
End Select
End Sub

Sub Process()          '處理主程式
ClearQueryLog (Me.Name) 'Add By Sindy 2010/01/22 清除查詢印表記錄檔欄位

Screen.MousePointer = vbHourglass
cnnConnection.Execute "delete from R050309 WHERE ID='" & strUserNum & "' "
strSQL1 = ""
strSQL2 = ""
'組字串
'系統類別
If Len(txt1(0)) <> 0 Then
   strSQL1 = strSQL1 & " and cp01 in (" & SQLGrpStr(txt1(0), 1) & ") "
   strSQL2 = strSQL2 & " and cp01 in (" & SQLGrpStr(txt1(0), 5) & ") "
   pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) 'Add By Sindy 2010/01/22
End If
'發文日
If Len(Trim(txt1(1))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
    strSQL2 = strSQL2 + " AND CP27>=" & Val(ChangeTStringToWString(txt1(1))) & " "
End If
If Len(Trim(txt1(2))) <> 0 Then
   strSQL1 = strSQL1 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
   strSQL2 = strSQL2 & " AND CP27<=" & Val(ChangeTStringToWString(txt1(2))) & " "
End If
pub_QL05 = pub_QL05 & ";" & Label2 & txt1(1) & "-" & txt1(2) 'Add By Sindy 2010/01/22
'申請國家
If Len(Trim(txt1(3))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)>='" & txt1(3) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)>='" & txt1(3) & "' "
End If
If Len(Trim(txt1(4))) <> 0 Then
    strSQL1 = strSQL1 + " AND SUBSTR(PA09,1,3)<='" & txt1(4) & "' "
    strSQL2 = strSQL2 + " AND SUBSTR(SP09,1,3)<='" & txt1(4) & "' "
End If
If Len(Trim(txt1(3))) <> 0 Or Len(Trim(txt1(4))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label3 & txt1(3) & "-" & txt1(4) 'Add By Sindy 2010/01/22
End If
'案件性質
If Len(Trim(txt1(5))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10>='" & txt1(5) & "' "
    strSQL2 = strSQL2 + " AND CP10>='" & txt1(5) & "' "
End If
If Len(Trim(txt1(6))) <> 0 Then
    strSQL1 = strSQL1 + " AND CP10<='" & txt1(6) & "' "
    strSQL2 = strSQL2 + " AND CP10<='" & txt1(6) & "' "
End If
If Len(Trim(txt1(5))) <> 0 Or Len(Trim(txt1(6))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label4 & txt1(5) & "-" & txt1(6) 'Add By Sindy 2010/01/22
End If
'承辦人
If Len(txt1(8)) <> 0 Then
    strSQL1 = strSQL1 + " AND CP14='" & txt1(8) & "' "
    strSQL2 = strSQL2 + " AND CP14='" & txt1(8) & "' "
    pub_QL05 = pub_QL05 & ";" & Label7 & txt1(8) 'Add By Sindy 2010/01/22
End If
'Modify By Cheng 2003/03/28
'搜尋資料時, 取消此條件
''是否計算多國案件
'If UCase(txt1(9)) = "Y" Then
'    'Modify By Cheng 2002/12/06
'    '含多國案件
''    strSQL1 = strSQL1 + " AND CP21='Y' "
''    strSQL2 = strSQL2 + " AND CP21='Y' "
'Else
'    strSQL1 = strSQL1 + " AND CP21 is null "
'    strSQL2 = strSQL2 + " AND CP21 is null "
'End If
''Add By Cheng 2003/04/09
''若不含不計件案件
'If Me.txt1(9).Text <> "" Then
'    strSQL1 = strSQL1 + " AND CP26 Is Null "
'    strSQL2 = strSQL2 + " AND CP26 Is Null "
'End If
'Add By Cheng 2002/12/05
strSQL1 = strSQL1 + " And CP04 ='00' "
strSQL2 = strSQL2 + " And CP04 ='00' "
'申請人
If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) <> 0 Then
    strSQL1 = strSQL1 + " AND ((PA26>='" & GetNewFagent(txt1(11)) & "' AND PA26<='" & GetNewFagent(txt1(12)) & "') OR (PA27>='" & GetNewFagent(txt1(11)) & "' AND PA27<='" & GetNewFagent(txt1(12)) & "') OR (PA28>='" & GetNewFagent(txt1(11)) & "' AND PA28<='" & GetNewFagent(txt1(12)) & "') OR (PA29>='" & GetNewFagent(txt1(11)) & "' AND PA29<='" & GetNewFagent(txt1(12)) & "') OR (PA30>='" & GetNewFagent(txt1(11)) & "' AND PA30<='" & GetNewFagent(txt1(12)) & "')) "
    strSQL2 = strSQL2 + " AND ((SP08>='" & GetNewFagent(txt1(11)) & "' AND SP08<='" & GetNewFagent(txt1(12)) & "') OR (SP58<='" & GetNewFagent(txt1(11)) & "' AND SP58<='" & GetNewFagent(txt1(12)) & "') OR (SP59>='" & GetNewFagent(txt1(11)) & "' AND SP59<='" & GetNewFagent(txt1(12)) & "')) "
Else
    If Len(Trim(txt1(11))) <> 0 And Len(Trim(txt1(12))) = 0 Then
        strSQL1 = strSQL1 + " AND (PA26>='" & GetNewFagent(txt1(11)) & "' OR PA27>='" & GetNewFagent(txt1(11)) & "' OR PA28>='" & GetNewFagent(txt1(11)) & "' OR PA29>='" & GetNewFagent(txt1(11)) & "' OR PA30>='" & GetNewFagent(txt1(11)) & "') "
        strSQL2 = strSQL2 + " AND (SP08>='" & GetNewFagent(txt1(11)) & "' OR SP58>='" & GetNewFagent(txt1(11)) & "' OR SP59>='" & GetNewFagent(txt1(11)) & "') "
    Else
        If Len(Trim(txt1(11))) = 0 And Len(Trim(txt1(12))) <> 0 Then
            strSQL1 = strSQL1 + " AND (PA26<='" & GetNewFagent(txt1(12)) & "' OR PA27<='" & GetNewFagent(txt1(12)) & "' OR PA28<='" & GetNewFagent(txt1(12)) & "' OR PA29<='" & GetNewFagent(txt1(12)) & "' OR PA30<='" & GetNewFagent(txt1(12)) & "') "
            strSQL2 = strSQL2 + " AND (SP08<='" & GetNewFagent(txt1(12)) & "' OR SP58<='" & GetNewFagent(txt1(12)) & "' OR SP59<='" & GetNewFagent(txt1(12)) & "') "
        End If
    End If
End If
If Len(Trim(txt1(11))) <> 0 Or Len(Trim(txt1(12))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label13 & txt1(11) & "-" & txt1(12) 'Add By Sindy 2010/01/22
End If
'代理人
If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) <> 0 Then
    '92.5.27 MODIFY BY SONIA
    'strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(13)) & "' AND PA75<='" & GetNewFagent(txt1(14)) & "' "
    'strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(13)) & "' AND SP26<='" & GetNewFagent(txt1(14)) & "' "
    strSQL1 = strSQL1 + " AND CP44>='" & GetNewFagent(txt1(13)) & "' AND CP44<='" & GetNewFagent(txt1(14)) & "' "
    strSQL2 = strSQL2 + " AND CP44>='" & GetNewFagent(txt1(13)) & "' AND CP44<='" & GetNewFagent(txt1(14)) & "' "
    '92.5.27 END
Else
    If Len(Trim(txt1(13))) <> 0 And Len(Trim(txt1(14))) = 0 Then
        '92.5.27 MODIFY BY SONIA
        'strSQL1 = strSQL1 + " AND PA75>='" & GetNewFagent(txt1(13)) & "' "
        'strSQL2 = strSQL2 + " AND SP26>='" & GetNewFagent(txt1(13)) & "' "
        strSQL1 = strSQL1 + " AND CP44>='" & GetNewFagent(txt1(13)) & "' "
        strSQL2 = strSQL2 + " AND CP44>='" & GetNewFagent(txt1(13)) & "' "
        '92.5.27 END
    Else
        If Len(Trim(txt1(13))) = 0 And Len(Trim(txt1(14))) <> 0 Then
            '92.5.27 MODIFY BY SONIA
            'strSQL1 = strSQL1 + " AND PA75<='" & GetNewFagent(txt1(14)) & "' "
            'strSQL2 = strSQL2 + " AND SP26<='" & GetNewFagent(txt1(14)) & "' "
            strSQL1 = strSQL1 + " AND CP44<='" & GetNewFagent(txt1(14)) & "' "
            strSQL2 = strSQL2 + " AND CP44<='" & GetNewFagent(txt1(14)) & "' "
            '92.5.27 END
        End If
    End If
End If
If Len(Trim(txt1(13))) <> 0 Or Len(Trim(txt1(14))) <> 0 Then
   pub_QL05 = pub_QL05 & ";" & Label14 & txt1(13) & "-" & txt1(14) 'Add By Sindy 2010/01/22
End If
'組合
'StrSQL = "SELECT ST02,CP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,decode(pa09,'000',ptm03,ptm04),decode(pa09,'000',cpm03,cpm04),NVL(NA03,NA04),CP05,CP18,PA75 AS B FROM CASEPROGRESS,STAFF,CASEPROPERTYMAP,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND CP14=ST01(+)  AND PTM01=1 AND PA08=PTM02(+) AND cp01=cpm01(+) AND cp10=cpm02(+) "
'StrSQL = StrSQL + " union all select ST02,CP27,CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',decode(sp09,'000',cpm03,cpm04),NVL(NA03,NA04),CP05,CP18,SP26 AS B FROM CASEPROGRESS,STAFF,CASEPROPERTYMAP,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP14=ST01(+)  AND cp01=cpm01(+) AND cp10=cpm02(+) "
'列印順序
If Val(txt1(7)) = 1 Then '依承辦人列印
   'Modify By Cheng 2002/09/17
   '剔除部門別為"P12"或承辦人代號為"72006"的資料
'    strSQL = "SELECT cp14," & SQLDate("cP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND CP14=ST01(+)  AND '1'=PTM01(+) AND PA08=PTM02(+) " & strsql1
'    strSQL = strSQL + " union all select cp14," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP14=ST01(+)  " & strsql2
   '92.1.13 MODIFY BY SONIA
   ' strSQL = "SELECT cp14," & SQLDate("cP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND CP14=ST01(+)  AND '1'=PTM01(+) AND PA08=PTM02(+) AND (CP14<>'72006' AND ST03<>'P12') " & strSQL1
   ' strSQL = strSQL + " union all select cp14," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP14=ST01(+) AND (CP14<>'72006' AND ST03<>'P12') " & strSQL2
'    strSQL = "SELECT cp14," & SQLDate("cP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND CP14=ST01(+)  AND '1'=PTM01(+) AND PA08=PTM02(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL1
'    strSQL = strSQL + " union all select cp14," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,STAFF,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP14=ST01(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL2
   '92.1.13 END
    'Modify By Cheng 2003/04/14
    '加欄位--是否算案件數
'    'Modify By Cheng 2003/03/28
'    '加欄位--是否多國案
'    strSQL = "SELECT cp14," & SQLDate("cP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21 FROM CASEPROGRESS,STAFF,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND CP14=ST01(+)  AND '1'=PTM01(+) AND PA08=PTM02(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL1
'    strSQL = strSQL + " union all select cp14," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21 FROM CASEPROGRESS,STAFF,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP14=ST01(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL2
    'strSQL = strSQL + " ORDER BY CP14,CP27,A "
    'modify by sonia 2016/3/3 +CP14<>'87025'
    strSql = "SELECT cp14," & SQLDate("cP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21, CP26 FROM CASEPROGRESS,STAFF,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND CP14=ST01(+)  AND '1'=PTM01(+) AND PA08=PTM02(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND CP14<>'87025' AND ST03<>'P12') " & strSQL1
    strSql = strSql + " union all select cp14," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21, CP26 FROM CASEPROGRESS,STAFF,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP14=ST01(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND CP14<>'87025' AND ST03<>'P12') " & strSQL2
Else '依代理人列印
   'Modify By Cheng 2002/09/17
   '剔除部門別為"P12"或承辦人代號為"72006"的資料
'    strSQL = "SELECT CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) " & strsql1
'    strSQL = strSQL + " union all select CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) " & strsql2
    '92.1.12 MODIFY BY SONIA
    'strSQL = "SELECT CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND (CP14<>'72006' AND ST03<>'P12') " & strSQL1
    'strSQL = strSQL + " union all select CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND (CP14<>'72006' AND ST03<>'P12') " & strSQL2
'    strSQL = "SELECT CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL1
'    strSQL = strSQL + " union all select CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "' FROM CASEPROGRESS,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL2
    '92.1.12 END
    'Modify By Cheng 2003/04/14
    '加欄位--是否算案件數
'    'Modify By Cheng 2003/03/28
'    '加欄位--是否多國案
'    strSQL = "SELECT CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21 FROM CASEPROGRESS,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL1
'    strSQL = strSQL + " union all select CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21 FROM CASEPROGRESS,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND ST03<>'P12') " & strSQL2
    'strSQL = strSQL + " ORDER BY B,CP27,A "
    'modify by sonia 2016/3/3 +CP14<>'87025'
    strSql = "SELECT CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(PA05,NVL(PA06,PA07)),PA11,ptm03,CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21, CP26 FROM CASEPROGRESS,STAFF,NATION,PATENT,PATENTTRADEMARKMAP WHERE cp01=PA01(+) AND cp02=PA02(+) AND cp03=PA03(+) AND cp04=PA04(+) AND pa09=na01(+) AND CP14=ST01(+) AND '1'=PTM01(+) AND PA08=PTM02(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND CP14<>'87025' AND ST03<>'P12') " & strSQL1
    strSql = strSql + " union all select CP44," & SQLDate("CP27") & ",CP01||'-'||CP02||'-'||CP03||'-'||CP04 AS A,NVL(SP05,NVL(SP06,SP07)),SP11,'',CP10,NVL(NA03,NA04)," & SQLDate("CP05") & ",CP18,'" & strUserNum & "', CP21, CP26 FROM CASEPROGRESS,STAFF,NATION,SERVICEPRACTICE WHERE cp01=SP01(+) AND cp02=SP02(+) AND cp03=SP03(+) AND cp04=SP04(+) AND sp09=na01(+) AND CP14=ST01(+) AND CP10 NOT IN ('907','913') AND (CP14<>'72006' AND CP14<>'87025' AND ST03<>'P12') " & strSQL2
End If
pub_QL05 = pub_QL05 & ";" & Label5 & txt1(7) & Label6  'Add By Sindy 2010/01/22
'是否列印明細
If txt1(10) = "Y" Or txt1(10) = "y" Then
   pub_QL05 = pub_QL05 & ";" & Label11 & txt1(10) 'Add By Sindy 2010/01/22
End If
cnnConnection.Execute "insert into r050309 " & strSql
CheckOC
strSql = "select * From r050309 where id='" & strUserNum & "' "
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
   InsertQueryLog (adoRecordset.RecordCount) 'Add By Sindy 2010/01/22
Else
   InsertQueryLog (0) 'Add By Sindy 2010/01/22
   ShowNoData
   Screen.MousePointer = vbDefault
   Exit Sub
End If
CheckOC
'是否列印明細
If txt1(10) = "Y" Or txt1(10) = "y" Then
    PrintData           '全印
Else
    PrintData1          '印小計與合計
End If

Screen.MousePointer = vbDefault
End Sub

Sub PrintData() '全印(含明細)
If Val(txt1(7)) = 1 Then '依承辦人分組
   strSql = "SELECT DISTINCT R009001,ST02 FROM R050309,STAFF WHERE ID='" & strUserNum & "' AND R009001=ST01(+) GROUP BY R009001,ST02 "
Else '依代理人分組
   strSql = "SELECT DISTINCT R009001,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) FROM R050309,FAGENT WHERE ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) GROUP BY R009001,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) "
End If

CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        StrTemp6 = CheckStr(adoRecordset.Fields(1))
        PrintTitle
        If Val(txt1(7)) = 1 Then
            If Len(CheckStr(adoRecordset.Fields(0))) = 0 Then
                'Modify By Cheng 2003/03/28
'               strSQL = "SELECT ST02,R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001 FROM R050309,STAFF,CASEPROPERTYMAP WHERE R009001 is null  AND ID='" & strUserNum & "' AND R009001=ST01(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  ORDER BY R009001,R009003,R009002 "
                'Modify By Cheng 2003/04/14
'               strSQL = "SELECT ST02,R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011 FROM R050309,STAFF,CASEPROPERTYMAP WHERE R009001 is null  AND ID='" & strUserNum & "' AND R009001=ST01(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  ORDER BY R009001,R009003,R009002 "
               strSql = "SELECT ST02,R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011,R009012 FROM R050309,STAFF,CASEPROPERTYMAP WHERE R009001 is null  AND ID='" & strUserNum & "' AND R009001=ST01(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  ORDER BY R009001,R009003,R009002 "
            Else
                'Modify By Cheng 2003/03/28
'               strSQL = "SELECT ST02,R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001 FROM R050309,STAFF,CASEPROPERTYMAP  WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND R009001=ST01(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
                'Modify By Cheng 2003/04/14
'               strSQL = "SELECT ST02,R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011 FROM R050309,STAFF,CASEPROPERTYMAP  WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND R009001=ST01(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
               strSql = "SELECT ST02,R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011,R009012 FROM R050309,STAFF,CASEPROPERTYMAP  WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND R009001=ST01(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
            End If
        Else
            If Len(CheckStr(adoRecordset.Fields(0))) = 0 Then
                'Modify By Cheng 2003/03/28
'               strSQL = "SELECT NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001 FROM R050309,FAGENT,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
                'Modify By Cheng 2003/04/14
'               strSQL = "SELECT NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011 FROM R050309,FAGENT,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
               strSql = "SELECT NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011,R009012 FROM R050309,FAGENT,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
            Else
                'Modify By Cheng 2003/03/28
'               strSQL = "SELECT NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001 FROM R050309,FAGENT,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
                'Modify By Cheng 2003/04/14
'               strSQL = "SELECT NVL(FA04,NVL(FA05||' '||FA63||' '||FA64||' '||FA65,FA06)),R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011 FROM R050309,FAGENT,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
               strSql = "SELECT NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)),R009002,R009003,R009004,R009005,R009006,nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009008,R009009,R009010,R009001,R009011,R009012 FROM R050309,FAGENT,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) AND 'CFP'=CPM01(+) AND R009007=CPM02(+) ORDER BY R009001,R009003,R009002 "
            End If
        End If
        CheckOC3
        AdoRecordSet3.CursorLocation = adUseClient
        AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If AdoRecordSet3.RecordCount <> 0 And AdoRecordSet3.RecordCount > 0 Then
            AdoRecordSet3.MoveFirst
            Do While AdoRecordSet3.EOF = False
                For i = 0 To 9
                    strTemp(i) = CheckStr(AdoRecordSet3.Fields(i))
                Next i
                strTemp(0) = StrConv(MidB(StrConv(strTemp(0), vbFromUnicode), 1, 24), vbUnicode)
                'Modify By Cheng 2003/04/14
                '更改案件名稱長度
'                strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 38), vbUnicode)
                strTemp(3) = StrConv(MidB(StrConv(strTemp(3), vbFromUnicode), 1, 30), vbUnicode)
                strTemp(4) = StrConv(MidB(StrConv(strTemp(4), vbFromUnicode), 1, 8), vbUnicode)
                strTemp(6) = StrConv(MidB(StrConv(strTemp(6), vbFromUnicode), 1, 8), vbUnicode)
                strTemp(10) = "" & AdoRecordSet3.Fields(11).Value '多國案
                strTemp(11) = "" & AdoRecordSet3.Fields(12).Value '計件案
                If iPrint > 10000 Then
                    PrintEnd
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle
                End If
                PrintDatil
               AdoRecordSet3.MoveNext
            Loop
        End If
        CheckOC3
        '計件合計
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        m_dblSubTotal = 0
        If Len(CheckStr(adoRecordset.Fields(0))) = 0 Then
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, r009012 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null And r009012 Is Null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, r009012 order by to_number(R009007) "
        Else
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, r009012 FROM R050309,CASEPROPERTYMAP WHERE ltrim(rtrim(R009001))='" & CheckStr(adoRecordset.Fields(0)) & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) And r009012 Is Null AND ID='" & strUserNum & "' " & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, r009012 order by to_number(R009007)  "
        End If
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            '若案件性質個數為5的倍數
            If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 = 0 Then
                adoRecordset1.MoveFirst
                For j = 1 To (adoRecordset1.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    If j = 1 Then PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
                    PrintTotil
                Next j
            '若案件性質個數非為5的倍數
            Else
                '若案件性質個數小於5個
                If adoRecordset1.RecordCount < 5 Then
                    adoRecordset1.MoveFirst
                    For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    For k = adoRecordset1.RecordCount Mod 5 To 4
                        StrTemp4(k) = ""
                        StrTemp5(k) = ""
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
                    PrintTotil
                '若案件性質個數超過5個
                Else
                    If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 <> 0 Then
                        adoRecordset1.MoveFirst
                        For j = 1 To (adoRecordset1.RecordCount \ 5)
                            For k = 0 To 4
                                StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                                StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                                If adoRecordset1.EOF = False Then
                                    adoRecordset1.MoveNext
                                End If
                            Next k
                            If iPrint > 10000 Then
                                PrintEnd
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle
                            End If
                            'Add By Cheng 2003/02/01
                            '在列印案件性質的第一列列印小計點數
                            If j = 1 Then PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
                            PrintTotil
                        Next j
                        For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                            StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                            StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                            If adoRecordset1.EOF = False Then
                                adoRecordset1.MoveNext
                            End If
                        Next k
                        For k = adoRecordset1.RecordCount Mod 5 To 4
                            StrTemp4(k) = ""
                            StrTemp5(k) = ""
                        Next k
                        If iPrint > 10000 Then
                            PrintEnd
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle
                        End If
                        PrintTotil
                    End If
                End If
            End If
            'Add By Cheng 2003/04/14
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print "計件合計：案件數 " & m_dblSubTotal & " 件"
            iPrint = iPrint + 300
        'Add By Cheng 2003/03/31
        '若無任何案件資料
        Else
            '列印小計點數
            PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
        End If
        CheckOC2
        
        'Modify By Cheng 2003/05/12
        '改成不計件合計
        '合計(含不計件)
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        If iPrint > 10000 Then
            PrintEnd
            Printer.NewPage
            Page = Page + 1
            PrintTitle
        End If
        m_dblSubTotal = 0
'        'Add By Cheng 2003/02/01
'        '列印小計點數
'        PrintCnt 0, adoRecordset.Fields(0).Value
        If Len(CheckStr(adoRecordset.Fields(0))) = 0 Then
            'Modify By Cheng 2003/04/14
            '判斷小計是否含不計件案
'            'Modify By Cheng 2003/03/28
'            '判斷多國案是否小計
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009011 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/05/07
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009012 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/05/12
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, R009012 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null And R009012='N' AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, R009012 order by to_number(R009007) "
        Else
            'Modify By Cheng 2003/04/14
            '判斷小計是否含不計件案
'            'Modify By Cheng 2003/03/28
'            '判斷多國案是否小計
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ltrim(rtrim(R009001))='" & CheckStr(adoRecordset.Fields(0)) & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) AND ID='" & strUserNum & "'  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007)  "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ltrim(rtrim(R009001))='" & CheckStr(adoRecordset.Fields(0)) & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) AND ID='" & strUserNum & "' " & IIf(Me.txt1(9).Text = "Y", "", " And R009011 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007)  "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ltrim(rtrim(R009001))='" & CheckStr(adoRecordset.Fields(0)) & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) AND ID='" & strUserNum & "'  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007)  "
            'Modify By Cheng 2003/05/07
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ltrim(rtrim(R009001))='" & CheckStr(adoRecordset.Fields(0)) & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) AND ID='" & strUserNum & "' " & IIf(Me.txt1(9).Text = "Y", "", " And R009012 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007)  "
            'Modify By Cheng 2003/05/12
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ltrim(rtrim(R009001))='" & CheckStr(adoRecordset.Fields(0)) & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) AND ID='" & strUserNum & "' " & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007)  "
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, R009012 FROM R050309,CASEPROPERTYMAP WHERE ltrim(rtrim(R009001))='" & CheckStr(adoRecordset.Fields(0)) & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) And R009012='N' AND ID='" & strUserNum & "' " & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, R009012 order by to_number(R009007)  "
        End If
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            '若案件性質個數為5的倍數
            If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 = 0 Then
                adoRecordset1.MoveFirst
                For j = 1 To (adoRecordset1.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    If j = 1 Then PrintCnt 0, adoRecordset.Fields(0).Value
                    PrintTotil
                Next j
            '若案件性質個數非為5的倍數
            Else
                '若案件性質個數小於5個
                If adoRecordset1.RecordCount < 5 Then
                    adoRecordset1.MoveFirst
                    For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    For k = adoRecordset1.RecordCount Mod 5 To 4
                        StrTemp4(k) = ""
                        StrTemp5(k) = ""
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    PrintCnt 0, "" & adoRecordset.Fields(0).Value
                    PrintTotil
                '若案件性質個數超過5個
                Else
                    If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 <> 0 Then
                        adoRecordset1.MoveFirst
                        For j = 1 To (adoRecordset1.RecordCount \ 5)
                            For k = 0 To 4
                                StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                                StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                                If adoRecordset1.EOF = False Then
                                    adoRecordset1.MoveNext
                                End If
                            Next k
                            If iPrint > 10000 Then
                                PrintEnd
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle
                            End If
                            'Add By Cheng 2003/02/01
                            '在列印案件性質的第一列列印小計點數
                            If j = 1 Then PrintCnt 0, adoRecordset.Fields(0).Value
                            PrintTotil
                        Next j
                        For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                            StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                            StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                            If adoRecordset1.EOF = False Then
                                adoRecordset1.MoveNext
                            End If
                        Next k
                        For k = adoRecordset1.RecordCount Mod 5 To 4
                            StrTemp4(k) = ""
                            StrTemp5(k) = ""
                        Next k
                        If iPrint > 10000 Then
                            PrintEnd
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle
                        End If
                        PrintTotil
                    End If
                End If
            End If
            'Add By Cheng 2003/04/14
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            'Modify By Cheng 2003/05/12
'            Printer.Print "合計(含不計件)：案件數 " & m_dblSubTotal & " 件"
            Printer.Print "不計件合計：案件數 " & m_dblSubTotal & " 件"
            iPrint = iPrint + 300
        'Add By Cheng 2003/03/31
        '若無任何案件資料
        Else
            '列印小計點數
            PrintCnt 0, adoRecordset.Fields(0).Value
        End If
        CheckOC2
        PrintEnd
        adoRecordset.MoveNext
        If adoRecordset.EOF = False Then
            Printer.NewPage
            Page = Page + 1
        End If
    Loop
Else
   Exit Sub
End If
CheckOC
StrTemp6 = "ALL"
PrintEnd
Printer.NewPage
Page = Page + 1
PrintTitle
'計件合計
m_dblSubTotal = 0
strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, r009012 FROM R050309,CASEPROPERTYMAP WHERE r009012 Is Null And ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, r009012 order by to_number(R009007) "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 = 0 Then
        adoRecordset.MoveFirst
        For j = 1 To (adoRecordset.RecordCount \ 5)
            For k = 0 To 4
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            If j = 1 Then PrintCnt_1 1, ""
            PrintTotil
        Next j
    Else
        If adoRecordset.RecordCount < 5 Then
            adoRecordset.MoveFirst
            For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            For k = adoRecordset.RecordCount Mod 5 To 4
                StrTemp4(k) = ""
                StrTemp5(k) = ""
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            PrintCnt_1 1, ""
            PrintTotil
        Else
            If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 <> 0 Then
                adoRecordset.MoveFirst
                For j = 1 To (adoRecordset.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                        If adoRecordset.EOF = False Then
                            adoRecordset.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '列印合計點數
                    If j = 1 Then PrintCnt_1 1, ""
                    PrintTotil
                Next j
                For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                For k = adoRecordset.RecordCount Mod 5 To 4
                    StrTemp4(k) = ""
                    StrTemp5(k) = ""
                Next k
                If iPrint > 10000 Then
                    PrintEnd
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle
                End If
                PrintTotil
            End If
        End If
    End If
    'Add By Cheng 2003/04/14
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "計件合計：案件數 " & m_dblSubTotal & " 件"
    iPrint = iPrint + 300
'Add By Cheng 2003/03/31
Else
    '列印合計點數
    PrintCnt_1 1, ""
End If
PrintEnd
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300

'Modify By Cheng 2003/05/12
'改成不計件合計
'合計(含不計件)
m_dblSubTotal = 0
''Add By Cheng 2003/02/01
''列印合計點數
'PrintCnt 1, ""
'Modify By Cheng 2003/04/14
'判斷小計是否含不計件案
''Modify By Cheng 2003/03/28
''判斷多國案是否小計
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009011 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'Modify By Cheng 2003/05/07
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009012 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'Modify By Cheng 2003/05/12
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, R009012 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' And R009012='N' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, R009012 order by to_number(R009007) "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 = 0 Then
        adoRecordset.MoveFirst
        For j = 1 To (adoRecordset.RecordCount \ 5)
            For k = 0 To 4
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            If j = 1 Then PrintCnt 1, ""
            PrintTotil
        Next j
    Else
        If adoRecordset.RecordCount < 5 Then
            adoRecordset.MoveFirst
            For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            For k = adoRecordset.RecordCount Mod 5 To 4
                StrTemp4(k) = ""
                StrTemp5(k) = ""
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            PrintCnt 1, ""
            PrintTotil
        Else
            If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 <> 0 Then
                adoRecordset.MoveFirst
                For j = 1 To (adoRecordset.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                        If adoRecordset.EOF = False Then
                            adoRecordset.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '列印合計點數
                    If j = 1 Then PrintCnt 1, ""
                    PrintTotil
                Next j
                For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                For k = adoRecordset.RecordCount Mod 5 To 4
                    StrTemp4(k) = ""
                    StrTemp5(k) = ""
                Next k
                If iPrint > 10000 Then
                    PrintEnd
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle
                End If
                PrintTotil
            End If
        End If
    End If
    'Add By Cheng 2003/04/14
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    'Modify By Cheng 2003/05/12
'    Printer.Print "合計(含不計件)：案件數 " & m_dblSubTotal & " 件"
    Printer.Print "不計件合計：案件數 " & m_dblSubTotal & " 件"
    iPrint = iPrint + 300
'Add By Cheng 2003/03/31
Else
    '列印合計點數
    PrintCnt 1, ""
End If
PrintEnd
CheckOC
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintDatil()           '印內容
For j = 1 To 8
    Printer.CurrentX = PLeft(j)
    Printer.CurrentY = iPrint
    Printer.Print strTemp(j)
Next j
'Add By Cheng 2003/03/28
'多國案
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print strTemp(10)
'Add By Cheng 2003/04/14
'計件案
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print strTemp(11)
'Modify By Cheng 2003/02/10
'設定點數的X座標
'Printer.CurrentX = PLeft(9) + 500 - Printer.TextWidth(strTemp(9))
Printer.CurrentX = PLeft(9) + Printer.TextWidth("點數") - Printer.TextWidth(Format(Val(strTemp(9)), "#.00"))
Printer.CurrentY = iPrint
'Modify By Cheng 2003/02/10
'設定點數格式
'Printer.Print strTemp(9)
Printer.Print Format(Val(strTemp(9)), "#.00")
iPrint = iPrint + 300
End Sub

Sub PrintData1()
If Val(txt1(7)) = 1 Then
   strSql = "SELECT DISTINCT R009001,ST02 FROM R050309,STAFF WHERE ID='" & strUserNum & "' AND R009001=ST01(+) GROUP BY R009001,ST02 "
Else
   strSql = "SELECT DISTINCT R009001,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) FROM R050309,FAGENT WHERE ID='" & strUserNum & "' AND SUBSTR(R009001,1,8)=FA01(+) AND SUBSTR(R009001,9,1)=FA02(+) GROUP BY R009001,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) "
End If
CheckOC
Page = 1
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    adoRecordset.MoveFirst
    Do While adoRecordset.EOF = False
        StrTemp6 = CheckStr(adoRecordset.Fields(1))
        PrintTitle
        '計件合計
        m_dblSubTotal = 0
        If Len(CheckStr(adoRecordset.Fields(0))) = 0 Then
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, r009012 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND r009012 Is Null And ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, r009012 order by to_number(R009007) "
        Else
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, r009012 FROM R050309,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' And r009012 Is Null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, r009012 order by to_number(R009007) "
        End If
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 = 0 Then
                adoRecordset1.MoveFirst
                For j = 1 To (adoRecordset1.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    If j = 1 Then PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
                    PrintTotil
                Next j
            Else
                If adoRecordset1.RecordCount < 5 Then
                    adoRecordset1.MoveFirst
                    For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    For k = adoRecordset1.RecordCount Mod 5 To 4
                        StrTemp4(k) = ""
                        StrTemp5(k) = ""
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
                    PrintTotil
                Else
                    If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 <> 0 Then
                        adoRecordset1.MoveFirst
                        For j = 1 To (adoRecordset1.RecordCount \ 5)
                            For k = 0 To 4
                                StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                                StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                                If adoRecordset1.EOF = False Then
                                    adoRecordset1.MoveNext
                                End If
                            Next k
                            If iPrint > 10000 Then
                                PrintEnd
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle
                            End If
                            'Add By Cheng 2003/02/01
                            '在列印案件性質的第一列列印小計點數
                            If j = 1 Then PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
                            PrintTotil
                        Next j
                        For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                            StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                            StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                            If adoRecordset1.EOF = False Then
                                adoRecordset1.MoveNext
                            End If
                        Next k
                        For k = adoRecordset1.RecordCount Mod 5 To 4
                            StrTemp4(k) = ""
                            StrTemp5(k) = ""
                        Next k
                        If iPrint > 10000 Then
                            PrintEnd
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle
                        End If
                        PrintTotil
                    End If
                End If
            End If
            'Add By Cheng 2003/04/14
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            Printer.Print "計件合計：案件數 " & m_dblSubTotal & " 件"
            iPrint = iPrint + 300
        'Add By Cheng 2003/03/31
        Else
            '列印小計點數
            PrintCnt_1 0, "" & adoRecordset.Fields(0).Value
        End If
        CheckOC2
        Printer.CurrentX = 500
        Printer.CurrentY = iPrint
        Printer.Print String(200, "-")
        iPrint = iPrint + 300
        'Modify By Cheng 2003/05/12
        '改成不計件合計
        '合計(含不計件)
        m_dblSubTotal = 0
        If Len(CheckStr(adoRecordset.Fields(0))) = 0 Then
            'strSQL = "SELECT R009007,COUNT(R009007) FROM R050309 WHERE R009001 is null AND ID='" & strUserNum & "'  GROUP BY R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/04/14
            '判斷小計是否含不計件案
'            'Modify By Cheng 2003/03/28
'            '判斷多國案是否小計
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009011 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/05/07
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009012 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/05/12
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, R009012 FROM R050309,CASEPROPERTYMAP WHERE R009001 is null AND ID='" & strUserNum & "' And R009012='N' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, R009012 order by to_number(R009007) "
        Else
            'strSQL = "SELECT R009007,COUNT(R009007) FROM R050309 WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' GROUP BY R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/04/14
            '判斷小計是否含不計件案
'            'Modify By Cheng 2003/03/28
'            '判斷多國案是否小計
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009011 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/05/07
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009012 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            'Modify By Cheng 2003/05/12
'            strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
            strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, R009012 FROM R050309,CASEPROPERTYMAP WHERE R009001='" & CheckStr(adoRecordset.Fields(0)) & "' AND ID='" & strUserNum & "' And R009012='N' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, R009012 order by to_number(R009007) "
        End If
        adoRecordset1.CursorLocation = adUseClient
        adoRecordset1.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
        If adoRecordset1.RecordCount <> 0 And adoRecordset1.RecordCount > 0 Then
            If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 = 0 Then
                adoRecordset1.MoveFirst
                For j = 1 To (adoRecordset1.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    If j = 1 Then PrintCnt 0, adoRecordset.Fields(0).Value
                    PrintTotil
                Next j
            Else
                If adoRecordset1.RecordCount < 5 Then
                    adoRecordset1.MoveFirst
                    For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                        StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                        If adoRecordset1.EOF = False Then
                            adoRecordset1.MoveNext
                        End If
                    Next k
                    For k = adoRecordset1.RecordCount Mod 5 To 4
                        StrTemp4(k) = ""
                        StrTemp5(k) = ""
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '在列印案件性質的第一列列印小計點數
                    PrintCnt 0, adoRecordset.Fields(0).Value
                    PrintTotil
                Else
                    If adoRecordset1.RecordCount \ 5 <> 0 And adoRecordset1.RecordCount Mod 5 <> 0 Then
                        adoRecordset1.MoveFirst
                        For j = 1 To (adoRecordset1.RecordCount \ 5)
                            For k = 0 To 4
                                StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                                StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                                If adoRecordset1.EOF = False Then
                                    adoRecordset1.MoveNext
                                End If
                            Next k
                            If iPrint > 10000 Then
                                PrintEnd
                                Printer.NewPage
                                Page = Page + 1
                                PrintTitle
                            End If
                            'Add By Cheng 2003/02/01
                            '在列印案件性質的第一列列印小計點數
                            If j = 1 Then PrintCnt 0, adoRecordset.Fields(0).Value
                            PrintTotil
                        Next j
                        For k = 0 To ((adoRecordset1.RecordCount Mod 5) - 1)
                            StrTemp4(k) = CheckStr(adoRecordset1.Fields(0))
                            StrTemp5(k) = CheckStr(adoRecordset1.Fields(1))
                            If adoRecordset1.EOF = False Then
                                adoRecordset1.MoveNext
                            End If
                        Next k
                        For k = adoRecordset1.RecordCount Mod 5 To 4
                            StrTemp4(k) = ""
                            StrTemp5(k) = ""
                        Next k
                        If iPrint > 10000 Then
                            PrintEnd
                            Printer.NewPage
                            Page = Page + 1
                            PrintTitle
                        End If
                        PrintTotil
                    End If
                End If
            End If
            'Add By Cheng 2003/04/14
            Printer.CurrentX = 500
            Printer.CurrentY = iPrint
            'Modify By Cheng 2003/05/12
'            Printer.Print "合計(含不計件)：案件數 " & m_dblSubTotal & " 件"
            Printer.Print "不計件合計：案件數 " & m_dblSubTotal & " 件"
            iPrint = iPrint + 300
        'Add By Cheng 2003/03/31
        Else
            '列印小計點數
            PrintCnt 0, adoRecordset.Fields(0).Value
        End If
        CheckOC2
        PrintEnd
        adoRecordset.MoveNext
        If adoRecordset.EOF = False Then
            Printer.NewPage
            Page = Page + 1
        End If
    Loop
Else
   Exit Sub
End If
CheckOC
Printer.NewPage
Page = Page + 1
StrTemp6 = "ALL"
PrintTitle
'計件合計
m_dblSubTotal = 0
strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, r009012 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' And r009012 Is Null AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, r009012 order by to_number(R009007) "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 = 0 Then
        adoRecordset.MoveFirst
        For j = 1 To (adoRecordset.RecordCount \ 5)
            For k = 0 To 4
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            If j = 1 Then PrintCnt_1 1, ""
            PrintTotil
        Next j
    Else
        If adoRecordset.RecordCount < 5 Then
            adoRecordset.MoveFirst
            For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            For k = adoRecordset.RecordCount Mod 5 To 4
                StrTemp4(k) = ""
                StrTemp5(k) = ""
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            PrintCnt_1 1, ""
            PrintTotil
        Else
            If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 <> 0 Then
                adoRecordset.MoveFirst
                For j = 1 To (adoRecordset.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                        If adoRecordset.EOF = False Then
                            adoRecordset.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '列印合計點數
                    If j = 1 Then PrintCnt_1 1, ""
                    PrintTotil
                Next j
                For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                For k = adoRecordset.RecordCount Mod 5 To 4
                    StrTemp4(k) = ""
                    StrTemp5(k) = ""
                Next k
                If iPrint > 10000 Then
                    PrintEnd
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle
                End If
                PrintTotil
            End If
        End If
    End If
    'Add By Cheng 2003/04/14
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    Printer.Print "計件合計：案件數 " & m_dblSubTotal & " 件"
    iPrint = iPrint + 300
'Add By Cheng 2003/03/31
Else
    '列印合計點數
    PrintCnt_1 1, ""
End If
PrintEnd
CheckOC
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
'Modify By Cheng 2003/05/12
'改成不計件合計
'合計(含不計件)
m_dblSubTotal = 0
'strSQL = "SELECT R009007,COUNT(R009007) FROM R050309 WHERE ID='" & strUserNum & "' and R009007 is not null  GROUP BY R009007 order  by to_number(R009007) "
'Modify By Cheng 2003/04/14
'判斷小計是否含不計件案
''Modify By Cheng 2003/03/28
''判斷多國案是否小計
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009011 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+)  GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'Modify By Cheng 2003/05/07
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) " & IIf(Me.txt1(9).Text = "Y", "", " And R009012 IS NULL ") & " GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
'Modify By Cheng 2003/05/12
'strSQL = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007 order by to_number(R009007) "
strSql = "SELECT nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),COUNT(nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007)),R009007, R009012 FROM R050309,CASEPROPERTYMAP WHERE ID='" & strUserNum & "' And R009012='N' AND 'CFP'=CPM01(+) AND R009007=CPM02(+) GROUP BY nvl(DECODE(R009009,'台灣',CPM03,CPM04),r009007),R009007, R009012 order by to_number(R009007) "
CheckOC
adoRecordset.CursorLocation = adUseClient
adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
    If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 = 0 Then
        adoRecordset.MoveFirst
        For j = 1 To (adoRecordset.RecordCount \ 5)
            For k = 0 To 4
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            If j = 1 Then PrintCnt 1, ""
            PrintTotil
        Next j
    Else
        If adoRecordset.RecordCount < 5 Then
            adoRecordset.MoveFirst
            For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                If adoRecordset.EOF = False Then
                    adoRecordset.MoveNext
                End If
            Next k
            For k = adoRecordset.RecordCount Mod 5 To 4
                StrTemp4(k) = ""
                StrTemp5(k) = ""
            Next k
            If iPrint > 10000 Then
                PrintEnd
                Printer.NewPage
                Page = Page + 1
                PrintTitle
            End If
            'Add By Cheng 2003/02/01
            '列印合計點數
            PrintCnt 1, ""
            PrintTotil
        Else
            If adoRecordset.RecordCount \ 5 <> 0 And adoRecordset.RecordCount Mod 5 <> 0 Then
                adoRecordset.MoveFirst
                For j = 1 To (adoRecordset.RecordCount \ 5)
                    For k = 0 To 4
                        StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                        StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                        If adoRecordset.EOF = False Then
                            adoRecordset.MoveNext
                        End If
                    Next k
                    If iPrint > 10000 Then
                        PrintEnd
                        Printer.NewPage
                        Page = Page + 1
                        PrintTitle
                    End If
                    'Add By Cheng 2003/02/01
                    '列印合計點數
                    If j = 1 Then PrintCnt 1, ""
                    PrintTotil
                Next j
                For k = 0 To ((adoRecordset.RecordCount Mod 5) - 1)
                    StrTemp4(k) = CheckStr(adoRecordset.Fields(0))
                    StrTemp5(k) = CheckStr(adoRecordset.Fields(1))
                    If adoRecordset.EOF = False Then
                        adoRecordset.MoveNext
                    End If
                Next k
                For k = adoRecordset.RecordCount Mod 5 To 4
                    StrTemp4(k) = ""
                    StrTemp5(k) = ""
                Next k
                If iPrint > 10000 Then
                    PrintEnd
                    Printer.NewPage
                    Page = Page + 1
                    PrintTitle
                End If
                PrintTotil
            End If
        End If
    End If
    'Add By Cheng 2003/04/14
    Printer.CurrentX = 500
    Printer.CurrentY = iPrint
    'Modify By Cheng 2003/05/12
'    Printer.Print "合計(含不計件)：案件數 " & m_dblSubTotal & " 件"
    Printer.Print "不計件合計：案件數 " & m_dblSubTotal & " 件"
    iPrint = iPrint + 300
'Add By Cheng 2003/03/31
Else
    '列印合計點數
    PrintCnt 1, ""
End If
PrintEnd
CheckOC
Printer.EndDoc
ShowPrintOk
End Sub

Sub PrintEnd()            '印結尾
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint
'Printer.Print String(200, "-")
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint + 300
'Printer.Print "註1.承辦人依案件性質小計且跳頁，再合計。"
'Printer.CurrentX = 500
'Printer.CurrentY = iPrint + 600
'Printer.Print "    2.本報表有兩種形式;另一報表的列印順序為代理人,且''承辦人''改為''代理人'',其餘相同."
'iPrint = iPrint + 300
End Sub

Sub PrintTotil()            '印小計
Dim k
For k = 0 To 4
    Printer.CurrentX = 500 + (k * 2000)
    Printer.CurrentY = iPrint
    Printer.Print StrConv(MidB(StrConv(StrTemp4(k), vbFromUnicode), 1, 10), vbUnicode)
    Printer.CurrentX = 2300 + (k * 2000) - Printer.TextWidth(StrTemp5(k))
    Printer.CurrentY = iPrint
    Printer.Print StrTemp5(k)
    'Add By Cheng 2003/04/14
    m_dblSubTotal = m_dblSubTotal + Val("0" & StrTemp5(k))
Next k
iPrint = iPrint + 300
End Sub

Sub PrintTitle()     '印抬頭
GetPleft
iPrint = 500
Printer.Orientation = 2
Printer.Font.Name = "細明體"
Printer.Font.Size = 22
Printer.Font.Bold = True
Printer.Font.Underline = True
Printer.CurrentX = 6300
Printer.CurrentY = iPrint
Printer.Print "承辦人發文明細表"
iPrint = iPrint + 500
Printer.Font.Size = 12
Printer.Font.Bold = False
Printer.Font.Underline = False
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
Printer.Print "發文日：" & Format(ChangeTStringToTDateString(txt1(1)) & " ", "@@@@@@@@@@") & "－" & ChangeTStringToTDateString(txt1(2))
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print "列印人：" & strUserName
'Add By Cheng 2003/04/09
Printer.CurrentX = 6500
Printer.CurrentY = iPrint
'Modify By Cheng 2003/04/14
'Printer.Print IIf(Me.txt1(9).Text = "N", "(不含不計件案件)", "(含不計件案件)")
'Printer.Print IIf(Me.txt1(9).Text = "Y", "(小計含不計件案件)", "(小計不含不計件案件)")
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "列印日期：" & Format(GetTaiwanTodayDate, "##/##/##")
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
If Val(txt1(7)) = 1 Then
    Printer.Print "承辦人：" & StrTemp6
Else
    Printer.Print "代理人：" & StrTemp6
End If
Printer.CurrentX = 13000
Printer.CurrentY = iPrint
Printer.Print "頁    次：" & str(Page)
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iPrint
Printer.Print "發文日"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iPrint
Printer.Print "本所案號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iPrint
Printer.Print "案件名稱"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iPrint
Printer.Print "申請案號"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iPrint
Printer.Print "專利種類"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iPrint
Printer.Print "案件性質"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iPrint
Printer.Print "申請國家"
Printer.CurrentX = PLeft(8)
Printer.CurrentY = iPrint
Printer.Print "收文日"
'Add By Cheng 2003/03/28
Printer.CurrentX = PLeft(10)
Printer.CurrentY = iPrint
Printer.Print "多國"
'Add By Cheng 2003/04/14
Printer.CurrentX = PLeft(11)
Printer.CurrentY = iPrint
Printer.Print "計件"
Printer.CurrentX = PLeft(9)
Printer.CurrentY = iPrint
Printer.Print "點數"
iPrint = iPrint + 300
Printer.CurrentX = 500
Printer.CurrentY = iPrint
Printer.Print String(200, "-")
iPrint = iPrint + 300
End Sub

Sub GetPleft()
Erase PLeft
PLeft(0) = 500
PLeft(1) = 500
PLeft(2) = 1500
PLeft(3) = 3600
PLeft(4) = 8300 - 1000
PLeft(5) = 9500 - 1000
PLeft(6) = 10600 - 1000
PLeft(7) = 11700 - 1000
PLeft(8) = 12800 - 1000
'Add By Cheng 2003/03/28
PLeft(10) = 14050 - 1000 '多國
'Add By Cheng 2003/03/28
PLeft(11) = 14050 '計件
'Modify By Cheng 2003/02/10
'PLeft(9) = 13800
'PLeft(9) = 13800 + 1000
PLeft(9) = 13800 + 1000 + 250 '點數
End Sub

Private Sub Form_Load()
MoveFormToCenter Me
txt1(0) = GetSystemKindByNick
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm050309 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   txt1(Index).SelStart = 0
   txt1(Index).SelLength = Len(txt1(Index))
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   'Add By Cheng 2002/09/16
   Select Case Index
   Case 7
      If KeyAscii <> 49 And KeyAscii <> 50 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 9 '小計是否含不計件案件
        'Modify By Cheng 2003/04/09
      If KeyAscii <> 89 And KeyAscii <> 8 Then
'      If KeyAscii <> 78 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   Case 10
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         KeyAscii = 0
      End If
   End Select
End Sub

Private Sub txt1_LostFocus(Index As Integer)
Select Case Index
Case 0
     strTemp1 = Split(UCase(GetSystemKindByNick), ",")
     strTemp2 = Split(UCase(txt1(0)), ",")
     For i = 0 To UBound(strTemp2)
        s = 0
        For j = 0 To UBound(strTemp1)
            If strTemp1(j) = strTemp2(i) Then
                s = 1
                Exit For
            End If
        Next j
        If s = 0 Then
            s = MsgBox(strUserName & " 沒有 " & strTemp2(i) & " 的權限 ", , "權限問題")
            txt1(0).SetFocus
            txt1(0).SelStart = 0
            txt1(0).SelLength = Len(txt1(0))
            Exit Sub
        End If
    Next i
Case 2, 4, 6
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
         txt1_GotFocus (Index - 1)
         Exit Sub
      End If
   Else
      blnClkSure = False
   End If
Case 7
   'Modify By Cheng 2002/09/26
   If Me.txt1(7).Text <> "" Then
      Select Case Val(txt1(Index))
      Case 1, 2
      Case Else
           s = MsgBox("列印順序只能 1 或 2 !!", , "USER 輸入錯誤")
           txt1(Index).SetFocus
           txt1(Index).SelStart = 0
           txt1(Index).SelLength = Len(txt1(Index))
           Exit Sub
      End Select
   End If
Case 8
     lbl1 = GetPrjSales(txt1(Index))
      'Add By Cheng 2002/09/26
      If Me.txt1(8).Text <> "" Then
         If Me.txt1(8).Text = Me.lbl1.Caption Then
            Me.lbl1.Caption = ""
            Me.txt1(8).SetFocus
            txt1_GotFocus 8
            Exit Sub
         End If
      End If
'Case 9 '小計是否含不計件案件
'   If Me.txt1(9).Text <> "" Then
'     Select Case txt1(Index)
'     Case "Y", ""
''     Case "N", ""
'     Case Else
''          s = MsgBox("小計是否計算多國案件只能為 Y !!", , "USER 輸入錯誤")
''          s = MsgBox("是否含不計件案件只能為 N 或 不輸 !!", , "USER 輸入錯誤")
'          s = MsgBox("小計是否含不計件案件只能為 Y 或 不輸 !!", , "USER 輸入錯誤")
'          txt1(Index).SetFocus
'          txt1(Index).SelStart = 0
'          txt1(Index).SelLength = Len(txt1(Index))
'          Exit Sub
'     End Select
'   End If
Case 10
   'Modify By Cheng 2002/09/26
   If Me.txt1(10).Text <> "" Then
     Select Case txt1(Index)
     Case "Y", "y", ""
     Case Else
          s = MsgBox("是否列印明細只能 Y !!", , "USER 輸入錯誤")
          txt1(Index).SetFocus
          txt1(Index).SelStart = 0
          txt1(Index).SelLength = Len(txt1(Index))
          Exit Sub
     End Select
   End If
Case 12
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If Len(txt1(Index - 1)) <> 0 Then
         If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
             s = MsgBox("申請人前 6 碼必須相同", , "USER 輸入錯誤")
             txt1(Index - 1).SetFocus
             Exit Sub
         End If
      End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
Case 14
   'Modify By Cheng 2002/09/16
   If blnClkSure = False Then
      If Len(txt1(Index - 1)) <> 0 Then
         If Left(txt1(Index - 1), 6) <> Left(txt1(Index), 6) Then
             s = MsgBox("代理人前 6 碼必須相同", , "USER 輸入錯誤")
             txt1(Index - 1).SetFocus
             Exit Sub
         End If
      End If
      If RunNick(txt1(Index - 1), txt1(Index)) Then
         txt1(Index - 1).SetFocus
      End If
   Else
      blnClkSure = False
   End If
End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
Select Case Index
Case 1, 2 '發文日起, 迄
   If PUB_CheckKeyInDate(Me.txt1(Index)) = -1 Then
      Cancel = True
      Me.txt1(Index).SetFocus
      txt1_GotFocus Index
   End If
End Select
End Sub

'Add By Cheng 2003/02/10
'列印點數小計或合計
Private Sub PrintCnt(intkind As Integer, strGroup As String)
'strCount : 1為計件合計, 2為不計件合計
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'小計
If intkind = 0 Then
'    strSQLA = "Select Sum(NVL(R009010,0)) From r050309 Where R009001='" & strGroup & "' And R009012 Is Null  And ID ='" & strUserNum & "' "
    StrSQLa = "Select Sum(NVL(R009010,0)) From r050309 Where R009001='" & strGroup & "' And R009012='N' And ID ='" & strUserNum & "' "
'合計
Else
'    strSQLA = "Select Sum(NVL(R009010,0)) From r050309 Where ID ='" & strUserNum & "' "
    StrSQLa = "Select Sum(NVL(R009010,0)) From r050309 Where ID ='" & strUserNum & "' And R009012='N' "
End If
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Printer.CurrentX = PLeft(9) + Printer.TextWidth("點數") - Printer.TextWidth(Format(rsA.Fields(0).Value, "#.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(rsA.Fields(0).Value, "#.00")
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

'Add By Cheng 2003/05/07
'列印點數小計或合計
Private Sub PrintCnt_1(intkind As Integer, strGroup As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

'小計
If intkind = 0 Then
    StrSQLa = "Select Sum(NVL(R009010,0)) From r050309 Where R009001='" & strGroup & "' And r009012 Is Null And ID ='" & strUserNum & "'  "
'合計
Else
    StrSQLa = "Select Sum(NVL(R009010,0)) From r050309 Where r009012 Is Null And ID ='" & strUserNum & "' "
End If
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Printer.CurrentX = PLeft(9) + Printer.TextWidth("點數") - Printer.TextWidth(Format(rsA.Fields(0).Value, "#.00"))
    Printer.CurrentY = iPrint
    Printer.Print Format(rsA.Fields(0).Value, "#.00")
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub
