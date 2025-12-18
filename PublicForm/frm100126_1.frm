VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100126_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "客戶重新委任案件查詢及列印"
   ClientHeight    =   3300
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4570
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4570
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd1 
      Height          =   3285
      Left            =   60
      TabIndex        =   20
      Top             =   3300
      Width           =   8955
      _ExtentX        =   15804
      _ExtentY        =   5786
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   7
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   7
      Text            =   "1"
      Top             =   2993
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   6
      Left            =   1410
      MaxLength       =   1
      TabIndex        =   5
      Text            =   "1"
      Top             =   1680
      Width           =   315
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   1410
      MaxLength       =   5
      TabIndex        =   6
      Text            =   "1,2,3"
      Top             =   2040
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1410
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1343
      Width           =   825
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   3
      Left            =   2190
      MaxLength       =   3
      TabIndex        =   3
      Top             =   990
      Width           =   645
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1410
      MaxLength       =   3
      TabIndex        =   2
      Top             =   990
      Width           =   585
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2820
      MaxLength       =   9
      TabIndex        =   1
      Top             =   630
      Width           =   1155
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1410
      MaxLength       =   9
      TabIndex        =   0
      Top             =   630
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   405
      Index           =   1
      Left            =   3630
      TabIndex        =   9
      Top             =   60
      Width           =   855
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Index           =   0
      Left            =   2730
      TabIndex        =   8
      Top             =   60
      Width           =   855
   End
   Begin MSForms.Label lblSalesName 
      Height          =   300
      Left            =   2310
      TabIndex        =   21
      Top             =   1350
      Width           =   1545
      Size            =   "2725;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Line Line2 
      BorderWidth     =   2
      X1              =   1770
      X2              =   2565
      Y1              =   1110
      Y2              =   1125
   End
   Begin VB.Line Line1 
      X1              =   2010
      X2              =   3360
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Label Label10 
      Caption         =   "1.已發文            2.待處理               3.不必重新委任"
      Height          =   555
      Left            =   1440
      TabIndex        =   19
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "(1.查詢  2.印表)"
      Height          =   180
      Left            =   1830
      TabIndex        =   18
      Top             =   3030
      Width           =   1200
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "(1.明細  2.統計)"
      Height          =   180
      Left            =   1830
      TabIndex        =   17
      Top             =   1710
      Width           =   1200
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "(可複選，請加,區隔)"
      Height          =   180
      Left            =   2340
      TabIndex        =   16
      Top             =   2070
      Width           =   1605
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "列印別："
      Height          =   180
      Left            =   660
      TabIndex        =   15
      Top             =   3030
      Width           =   720
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "資料顯示方式："
      Height          =   180
      Left            =   120
      TabIndex        =   14
      Top             =   1710
      Width           =   1260
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "重新委任類別："
      Height          =   180
      Left            =   120
      TabIndex        =   13
      Top             =   2070
      Width           =   1260
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "智權人員："
      Height          =   180
      Left            =   480
      TabIndex        =   12
      Top             =   1380
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "業務區："
      Height          =   180
      Left            =   660
      TabIndex        =   11
      Top             =   1027
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "客戶編號："
      Height          =   180
      Left            =   480
      TabIndex        =   10
      Top             =   675
      Width           =   900
   End
End
Attribute VB_Name = "frm100126_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/14 改成Form2.0 ; lblSalesName ; Printer列印未改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Public cmdState As Integer
Public Str100126SQL1 As String
Public Str100126SQL2 As String
Public Str100126SQL3 As String
Public StrAllCU As String
Public MyStrSQL  As String
Dim iLine As Integer
Dim MaxLine As Integer
Dim PLeft(7) As Integer
Dim strTemp(7) As String
Dim oPage As Integer
Dim SeekTemp(2) As String
Dim oFontH As Integer
'Added by Lydia 2019/11/01 利益衝突案件
Dim m_AllSys As String '預設全部系統別
Dim intCufaCnt As Integer '限閱案件X件
Dim SeColPA As String

Private Sub cmdok_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Sub PubShowNextData()
   Select Case cmdState
      Case 0
              cmdState = -1
              Screen.MousePointer = vbHourglass
              ClearQueryLog (Me.Name) 'Add By Sindy 2010/11/16 清除查詢印表記錄檔欄位
              DoEvents
              If ConstrainCheck = True Then
                  If doChgSqlStr Then
                      If txt1(7) = "1" Then
                          If fnSaveParentForm(Me) = False Then
                              Me.Enabled = True
                              Exit Sub
                          End If
                          pub_QL05 = pub_QL05 & ";" & Label6 & "1.查詢" 'Add By Sindy 2010/11/16
                          frm100126_2.Show
                      Else
                          pub_QL05 = pub_QL05 & ";" & Label6 & "2.印表" 'Add By Sindy 2010/11/16
                          PrintData
                      End If
                  End If
              End If
              Screen.MousePointer = vbDefault
      Case 1
              fnCloseAllFrm100
      Case Else
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm100126_1 = Nothing
End Sub

Private Sub txt1_Change(Index As Integer)
   Select Case Index
      '2007/7/16 ADD BY SONIA
      Case 0
               Me.txt1(1).Text = Me.txt1(Index).Text
      '2007/7/16 END
      Case 4
              If Len(txt1(Index)) > 4 Then
                 lblSalesName = GetStaffName(txt1(Index))
              Else
                 lblSalesName = ""
              End If
      Case Else
   End Select
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1, 2, 3, 4
              KeyAscii = UpperCase(KeyAscii)
      Case 5
              If (KeyAscii < Asc("1") Or KeyAscii > Asc("3")) And KeyAscii <> 8 And KeyAscii <> Asc(",") Then
                  KeyAscii = 0
                  Beep
              End If
      Case 6, 7
              If (KeyAscii < Asc("1") Or KeyAscii > Asc("2")) And KeyAscii <> 8 Then
                  KeyAscii = 0
                  Beep
              ElseIf KeyAscii = Asc("1") And Index = 6 Then
                  txt1(5).Enabled = True
              ElseIf KeyAscii = Asc("2") And Index = 6 Then
                  txt1(5).Enabled = False
                  txt1(5) = "1,2,3"
              End If
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
              If txt1(0) <> "" And txt1(1) <> "" Then
                  If Mid(txt1(0), 1, 6) <> Mid(txt1(1), 1, 6) Then
                      MsgBox "申請人前6碼必須相同！", vbExclamation
                      txt1(0).SetFocus
                      txt1_GotFocus 0
                      Cancel = True
                      Exit Sub
                  End If
              End If
      Case 3
              If txt1(2) <> "" And txt1(3) <> "" Then
                  If txt1(2) > txt1(3) Then
                      MsgBox "業務區範圍條件錯誤！", vbExclamation
                      txt1(2).SetFocus
                      txt1_GotFocus 2
                      Cancel = True
                      Exit Sub
                  End If
              End If
      Case 4
              If Len(txt1(Index)) > 4 Then
                 lblSalesName = GetStaffName(txt1(Index))
              Else
                 lblSalesName = ""
              End If
      Case Else
   End Select
End Sub
Private Function ConstrainCheck() As Boolean
   Dim bolCancel As Boolean
   ConstrainCheck = True
   If Trim(txt1(0)) = "" And Trim(txt1(1)) = "" And Trim(txt1(2)) = "" And Trim(txt1(3)) = "" And Trim(txt1(4)) = "" Then
      MsgBox "請最少輸入  客戶編號、業務區、智權人員  任一種！", vbExclamation
      txt1(0).SetFocus
      txt1_GotFocus 0
      ConstrainCheck = False
      Exit Function
   End If
   If Trim(txt1(5)) = "" And txt1(5).Enabled = True Then
      MsgBox "請輸入重新委任類別！", vbExclamation
      txt1(5).SetFocus
      txt1_GotFocus 5
      ConstrainCheck = False
      Exit Function
   End If
   If Trim(txt1(6)) = "" Then
      MsgBox "請輸入資料顯示方式！", vbExclamation
      txt1(6).SetFocus
      txt1_GotFocus 6
      ConstrainCheck = False
      Exit Function
   End If
    If txt1(0) <> "" Or txt1(1) <> "" Then
        If Mid(txt1(0), 1, 6) <> Mid(txt1(1), 1, 6) Then
            MsgBox "申請人前6碼必須相同！", vbExclamation
            txt1(0).SetFocus
            txt1_GotFocus 0
            ConstrainCheck = False
            Exit Function
        End If
    End If
    If txt1(2) <> "" Or txt1(3) <> "" Then
        If txt1(2) > txt1(3) Then
            MsgBox "業務區範圍條件錯誤！", vbExclamation
            txt1(2).SetFocus
            txt1_GotFocus 2
            ConstrainCheck = False
            Exit Function
        End If
        If txt1(2) = "" Then
            MsgBox "業務區範圍條件錯誤！", vbExclamation
            txt1(2).SetFocus
            txt1_GotFocus 2
            ConstrainCheck = False
            Exit Function
        End If
        If txt1(3) = "" Then
            MsgBox "業務區範圍條件錯誤！", vbExclamation
            txt1(3).SetFocus
            txt1_GotFocus 3
            ConstrainCheck = False
            Exit Function
        End If
        If txt1(6) = "2" Then txt1(5) = "1,2,3"
    End If

End Function

Function doChgSqlStr() As Boolean
   doChgSqlStr = False

Dim rsTmp As New ADODB.Recordset
Dim strFields(1 To 4) As String
Dim TheStrSql  As String, strTmp As String

   strTmp = ""
   '抓出所有申請人
   StrAllCU = ""
   MyStrSQL = ""
   If txt1(0) <> "" And txt1(1) <> "" Then
       MyStrSQL = MyStrSQL & " and cu01||cu02>='" & Mid(txt1(0) & "000000000", 1, 9) & "' and cu01||cu02<='" & Mid(txt1(1) & "000000000", 1, 9) & "' "
       pub_QL05 = pub_QL05 & ";" & Label1 & txt1(0) & "-" & txt1(1) 'Add By Sindy 2010/11/16
   End If
   If txt1(2) <> "" And txt1(3) <> "" Then
       MyStrSQL = MyStrSQL & " and cu12>='" & txt1(2) & "' and cu12<='" & txt1(3) & "' "
       pub_QL05 = pub_QL05 & ";" & Label2 & txt1(2) & "-" & txt1(3) 'Add By Sindy 2010/11/16
   End If
   If txt1(4) <> "" Then
       MyStrSQL = MyStrSQL & " and st02='" & txt1(4) & "' "
       pub_QL05 = pub_QL05 & ";" & Label3 & txt1(4) & lblSalesName 'Add By Sindy 2010/11/16
   End If
   Str100126SQL1 = ""
   Str100126SQL2 = ""
   Str100126SQL3 = ""
   If txt1(6) = "1" And txt1(7) = "1" Then
       TheStrSql = "  select distinct cu01||nvl(cu02,'0') from customer,staff  where cu13=st01(+) and cu01||cu02 in (select distinct pa26 from patent,caseprogress where " & _
                          "  cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa27 from patent,caseprogress where " & _
                          "  cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa28 from patent,caseprogress where " & _
                          "  cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa29 from patent,caseprogress where " & _
                          "  cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa30 from patent,caseprogress where " & _
                          "  cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa26 from patent,caseprogress where " & _
                          "  cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa27 from patent,caseprogress where " & _
                          "  cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa28 from patent,caseprogress where " & _
                          "  cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa29 from patent,caseprogress where " & _
                          "  cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) union select pa30 from patent,caseprogress where " & _
                          "  cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) ) " & MyStrSQL
       TheStrSql = TheStrSql & " order by 1 "
       If rsTmp.State = 1 Then rsTmp.Close
       rsTmp.CursorLocation = adUseClient
       rsTmp.Open TheStrSql, cnnConnection, adOpenStatic, adLockReadOnly
       If rsTmp.RecordCount > 0 Then
           rsTmp.MoveFirst
           Do While Not rsTmp.EOF
               StrAllCU = StrAllCU & CheckStr(rsTmp.Fields(0)) & ","
               rsTmp.MoveNext
           Loop
           StrAllCU = ChgNewStr(StrAllCU)
       Else
           InsertQueryLog (0) 'Add By Sindy 2010/11/16
           ShowNoData
           Exit Function
       End If
       rsTmp.Close
       '組所有資料ㄉ語法
       'Modified by Lydia 2019/11/01 利益衝突案件：加欄位
'            '已發文明細
'            strFields(1) = "decode(pa23,'1','','N')||pa01||'-'||pa02||'-'||pa03||'-'||pa04,pa11,nvl(pa05,nvl(pa06,pa07)),sqldatet(pa10),sqldatet(cp27),10,pa01||'-'||pa02||'-'||pa03||'-'||pa04 "
'            '待處理明細
'            strFields(2) = "decode(pa23,'1','','N')||pa01||'-'||pa02||'-'||pa03||'-'||pa04,pa11,nvl(pa05,nvl(pa06,pa07)),sqldatet(pa10),'',20,pa01||'-'||pa02||'-'||pa03||'-'||pa04 "
'            '不必重新委任明細
'            strFields(3) = "decode(pa23,'1','','N')||pa01||'-'||pa02||'-'||pa03||'-'||pa04,pa11,nvl(pa05,nvl(pa06,pa07)),sqldatet(pa10),sqldatet(cp27),30,pa01||'-'||pa02||'-'||pa03||'-'||pa04 "
            SeColPA = ",pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno"
            '已發文明細
            strFields(1) = "decode(pa23,'1','','N')||pa01||'-'||pa02||'-'||pa03||'-'||pa04 as 本所案號,pa11 as 申請案號,nvl(pa05,nvl(pa06,pa07)) as 案件名稱,sqldatet(pa10) as 申請日,sqldatet(cp27) as 重新委任發文日 ,10 as FSort ,pa01||'-'||pa02||'-'||pa03||'-'||pa04 as CaseNo" & SeColPA
            '待處理明細
            strFields(2) = "decode(pa23,'1','','N')||pa01||'-'||pa02||'-'||pa03||'-'||pa04 as 本所案號,pa11 as 申請案號,nvl(pa05,nvl(pa06,pa07)) as 案件名稱,sqldatet(pa10) as 申請日,'' as 重新委任發文日,20 as FSort ,pa01||'-'||pa02||'-'||pa03||'-'||pa04 as CaseNo " & SeColPA
            '不必重新委任明細
            strFields(3) = "decode(pa23,'1','','N')||pa01||'-'||pa02||'-'||pa03||'-'||pa04 as 本所案號,pa11 as 申請案號,nvl(pa05,nvl(pa06,pa07)) as 案件名稱,sqldatet(pa10) as 申請日,sqldatet(cp27) as 重新委任發文日,30 as FSort ,pa01||'-'||pa02||'-'||pa03||'-'||pa04 as CaseNo " & SeColPA
       'end 2019/11/01
       If InStr(1, txt1(5), "1") <> 0 Then
           strTmp = strTmp & IIf(strTmp = "", "", ";") & "1.已發文" 'Add By Sindy 2010/11/16
           Str100126SQL1 = "select " & strFields(1) & " from patent,caseprogress where CP10='928' AND pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp27 is not null and cp27<>19221111 "
       End If
       If InStr(1, txt1(5), "2") <> 0 Then
           strTmp = strTmp & IIf(strTmp = "", "", ";") & "2.待處理" 'Add By Sindy 2010/11/16
           Str100126SQL2 = "select " & strFields(2) & " from patent,caseprogress where CP10='928' AND pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp27 is null  "
       End If
       If InStr(1, txt1(5), "3") <> 0 Then
           strTmp = strTmp & IIf(strTmp = "", "", ";") & "3.不必重新委任" 'Add By Sindy 2010/11/16
           Str100126SQL3 = "select " & strFields(3) & " from patent,caseprogress where CP10='928' AND pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp27=19221111 "
       End If
       If strTmp <> "" Then pub_QL05 = pub_QL05 & ";" & Label4 & strTmp 'Add By Sindy 2010/11/16
   End If
   doChgSqlStr = True
End Function

Sub PrintData()
Dim rsTmp As New ADODB.Recordset
Dim strSql  As String
'Added by Lydia 2019/11/01 利益衝突案件：於後面增加欄位
Dim strMid As String, strGrp As String
Dim strJumpList As String '已排除的本所案號

   strSql = ""
   'Added by Lydia 2019/11/01利益衝突案件：於後面增加欄位
   SeColPA = " ,pa26 as cust01,pa27 as cust02,pa28 as cust03,pa29 as cust04,pa30 as cust05,pa75 as fcno "
   intCufaCnt = 0
   m_AllSys = "FCP,P"
   'end 2019/11/01

   If txt1(6) = "1" Then
      pub_QL05 = pub_QL05 & ";" & Label5 & "1.明細" 'Add By Sindy 2010/11/16
      If InStr(1, txt1(5), "1") <> 0 Then
          'Modified by Lydia 2019/11/01 利益衝突案件：加欄位
'          strSql = strSql & " select pa26,cp01||'-'||cp02||'-'||cp03||'-'||cp04,pa11,nvl(pa05,nvl(pa06,pa07)),sqldatet(pa10),'10' from customer, "
'          strSql = strSql & " (select distinct pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB"
          strSql = strSql & " select pa26,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as CaseNo,pa11,nvl(pa05,nvl(pa06,pa07)) as CaseName,sqldatet(pa10) as pa10,'10' as Fsort, cust01, cust02, cust03, cust04, cust05, fcno from customer, "
          strSql = strSql & " (select distinct pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is not null and cp27<>19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB"
          'end 2019/11/01
          strSql = strSql & " where cu01||cu02=TmpTB.pa26 " & MyStrSQL
      End If
      If InStr(1, txt1(5), "2") <> 0 Then
          If strSql <> "" Then
              strSql = strSql & " union "
          End If
          'Modified by Lydia 2019/11/01 利益衝突案件：加欄位
'          strSql = strSql & " select pa26,cp01||'-'||cp02||'-'||cp03||'-'||cp04,pa11,nvl(pa05,nvl(pa06,pa07)),sqldatet(pa10),'20' from customer, "
'          strSql = strSql & " (select distinct pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB"
          strSql = strSql & " select pa26,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as CaseNo,pa11,nvl(pa05,nvl(pa06,pa07)) as CaseName,sqldatet(pa10) as pa10,'20' as Fsort, cust01, cust02, cust03, cust04, cust05, fcno from customer, "
          strSql = strSql & " (select distinct pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27 is null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB"
          'end 2019/11/01
          strSql = strSql & " where cu01||cu02=TmpTB.pa26 " & MyStrSQL
      End If
      If InStr(1, txt1(5), "3") <> 0 Then
          If strSql <> "" Then
              strSql = strSql & " union "
          End If
          'Modified by Lydia 2019/11/01 利益衝突案件：加欄位
'          strSql = strSql & " select pa26,cp01||'-'||cp02||'-'||cp03||'-'||cp04,pa11,nvl(pa05,nvl(pa06,pa07)),sqldatet(pa10),'30' from customer, "
'          strSql = strSql & " (select distinct pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
'          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB"
          strSql = strSql & " select pa26,cp01||'-'||cp02||'-'||cp03||'-'||cp04 as CaseNo,pa11,nvl(pa05,nvl(pa06,pa07)) as CaseName,sqldatet(pa10) as PA10,'30' as Fsort, cust01, cust02, cust03, cust04, cust05, fcno from customer, "
          strSql = strSql & " (select distinct pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa26,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa27,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa28,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa29,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)"
          strSql = strSql & " union select pa30,cp27,cp09,cp01,cp02,cp03,cp04,pa11,pa10,pa05,pa06,pa07 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp27=19221111 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB"
      'end 2019/11/01
          strSql = strSql & " where cu01||cu02=TmpTB.pa26 " & MyStrSQL
      End If
      strSql = strSql & " order by 1,6,2 "
      Call ProcDataByCase("1", strSql) 'Added by Lydia 2019/11/01
   Else
      pub_QL05 = pub_QL05 & ";" & Label5 & "2.統計" 'Add By Sindy 2010/11/16
      SeColPA = ",pa01||'-'||pa02||'-'||pa03||'-'||pa04 as caseno" & SeColPA 'Added by Lydia 2019/11/01
      strSql = "select a0902 as oA1,st02 as oA2,cu01||cu02 as oA3,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as oA4 ,"
      strSql = strSql & " sum(decode(cp27,null,0,19221111,0,1)) as oA5,sum(decode(cp27,null,1,19221111,0,0)) as oA6,sum(decode(cp27,null,0,19221111,1,0)) as oA7,cu13 "
      strSql = strSql & ",caseno , cust01, cust02, cust03, cust04, cust05, fcno " 'Added by Lydia 2019/11/01 利益衝突案件：加欄位
      'Modified by Lydia 2019/11/01 利益衝突案件：加欄位SeColPA
      strSql = strSql & " from customer,staff,acc090,(select distinct pa26,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa27,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa28,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa29,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa30,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='P' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa26,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa27,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa28,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa29,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      strSql = strSql & " union select pa30,cp27,cp09 " & SeColPA & " from patent,caseprogress where cp01='FCP' and cp10='928' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) )  TmpTB "
      'end 2019/11/01
      strSql = strSql & " where cu01||cu02=TmpTB.pa26 and cu13=st01(+) and cu12=a0901(+) " & frm100126_1.MyStrSQL
      'Modified by Lydia 2019/11/01 2019/11/01 利益衝突案件：加欄位
      'strSql = strSql & " group by a0902,st02,cu01||cu02,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu13 order by 8,1,2,3 "
      strSql = strSql & " group by a0902,st02,cu01||cu02,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)),cu13,caseno , cust01, cust02, cust03, cust04, cust05, fcno "
      Call ProcDataByCase("2", strSql) 'Added by Lydia 2019/11/01
   End If
   
'Remove by Lydia 2019/11/01 改成模組ProcDataByCase
'   If rsTmp.State = 1 Then rsTmp.Close
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'       InsertQueryLog (rsTmp.RecordCount) 'Add By Sindy 2010/11/16
'       Set grd1.Recordset = rsTmp
'       MaxLine = 53
'       oFontH = 300
'       If txt1(6) = "1" Then
'           PrintData1
'       Else
'           PrintData2
'       End If
'   Else
'       InsertQueryLog (0) 'Add By Sindy 2010/11/16
'       ShowNoData
'   End If
End Sub

'Added by Lydia 2019/11/01 利益衝突案件：逐案號判斷
Private Sub ProcDataByCase(ByVal pType As String, ByRef pSQL As String)
Dim rsAD As New ADODB.Recordset
Dim strMid As String, strGrp As String
Dim strJumpList As String '已排除的本所案號
Dim intJump As Integer '空白列數
Dim mESeqNo As String '暫存TB編號
Dim strA1 As String
Dim dblRow As Double 'Add By Sindy 2025/9/3

    rsAD.CursorLocation = adUseClient
    rsAD.Open pSQL, cnnConnection, adOpenDynamic, adLockBatchOptimistic
    If rsAD.RecordCount > 0 Then
        dblRow = rsAD.RecordCount 'Add By Sindy 2025/9/3

        If pType = "2" Then '統計=>將明細資料丟到暫存檔rdatafactory
            Set RsTemp = PUB_CreateRecordset(rsAD, , , , Me.Name, mESeqNo)
        End If
        If strSrvDate(1) >= XY特殊權限啟用日 And XY特殊權限範圍 <> "" Then
            rsAD.MoveFirst
            Do While rsAD.EOF = False
                strMid = "" & rsAD.Fields("CaseNo")
                '利益衝突案件：逐案號判斷
                If Len(strMid) > 9 Then
                    If strJumpList <> "" And InStr(strJumpList, strMid) > 0 Then
                        '剔除重複的本所案號
                        rsAD.Delete
                    Else
                        If strGrp <> strMid Then
                            If PUB_ChkCufaByCase(Me.Name, m_AllSys, strMid, "" & rsAD.Fields("cust01") & "," & rsAD.Fields("cust02") & "," & rsAD.Fields("cust03") & "," & rsAD.Fields("cust04") & "," & rsAD.Fields("cust05"), "" & rsAD.Fields("fcno")) = False Then
                                strJumpList = strJumpList & strMid & ","
                                intCufaCnt = intCufaCnt + 1
                                rsAD.Delete
                                If pType = "2" Then '統計
                                   strA1 = "delete from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and r009='" & strMid & "' "
                                   cnnConnection.Execute strA1
                                End If
                            End If
                        End If
                    End If
                End If
                strGrp = strMid
                rsAD.MoveNext
            Loop
            '利益衝突案件：限閱案件
            If intCufaCnt > 0 Then
                pub_QL05 = pub_QL05 & "(含限閱" & intCufaCnt & "筆)" 'Add By Sindy 2025/9/3
                MsgBox MsgText(1109) & " " & intCufaCnt & " 件", vbInformation, MsgText(1110)
            End If
            InsertQueryLog (dblRow)
            If rsAD.RecordCount = 0 Then
                  GoTo JumpToNoData
            End If
        Else
            InsertQueryLog (rsAD.RecordCount)
        End If
        If pType = "2" Then '統計
            If rsAD.State <> adStateClosed Then rsAD.Close
            rsAD.CursorLocation = adUseClient
            strA1 = "select R001 as OA1,R002 as OA2,R003 as OA3,R004 as OA4,SUM(R005) as OA5,SUM(R006) as OA6,SUM(R007) as OA7,R008 as CU13 from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "'"
            strA1 = strA1 & " group by r001,r002,r003,r004,r008 order by 8,1,2,3 "
            rsAD.Open strA1, cnnConnection, adOpenStatic, adLockReadOnly
            If rsAD.RecordCount = 0 Then
                GoTo JumpToNoData
            End If
        End If
        
        Set GRD1.Recordset = rsAD
        MaxLine = 53
        oFontH = 300
        If pType = "1" Then
            PrintData1
        Else
            PrintData2
        End If
    Else
        InsertQueryLog (0)
JumpToNoData:
        ShowNoData
    End If
    Set rsAD = Nothing
End Sub

Sub PrintData1()
Dim i As Long
Dim oCount As Integer
On Error GoTo 0

   If Printer.Orientation <> 1 Then
       Printer.Orientation = 1
   End If
   GetPleft
   SeekTemp(1) = ""
   SeekTemp(2) = ""
   oCount = 0
   PrintT GRD1.TextMatrix(1, 0)
   With GRD1
       For i = 1 To .Rows - 1
           If SeekTemp(1) <> .TextMatrix(i, 0) Then
               If SeekTemp(1) <> "" Then
                   Printer.CurrentX = 0
                   Printer.CurrentY = iLine * oFontH
                   Printer.Print String(200, "-")
                   Printer.CurrentX = 0
                   Printer.CurrentY = (iLine + 1) * oFontH
                   Printer.Print "共 " & oCount & " 筆"
                   oCount = 0
                   Printer.NewPage
               End If
               PrintT .TextMatrix(i, 0)
               SeekTemp(1) = .TextMatrix(i, 0)
               SeekTemp(2) = .TextMatrix(i, 5)
               PrintT1 Trim(.TextMatrix(i, 5)), .TextMatrix(i, 0)
           End If
           If SeekTemp(2) <> .TextMatrix(i, 5) Then
               Printer.CurrentX = 0
               Printer.CurrentY = iLine * oFontH
               Printer.Print String(200, "-")
               Printer.CurrentX = 0
               Printer.CurrentY = (iLine + 1) * oFontH
               Printer.Print "共 " & oCount & " 筆"
               oCount = 0
               iLine = iLine + 3
               SeekTemp(2) = .TextMatrix(i, 5)
               PrintT1 Trim(.TextMatrix(i, 5)), .TextMatrix(i, 0)
           End If
           If iLine > MaxLine Then
               Printer.NewPage
               PrintT .TextMatrix(i, 0)
               PrintT1 Trim(.TextMatrix(i, 5)), .TextMatrix(i, 0)
           End If
           '本所案號
           Printer.CurrentX = PLeft(1)
           Printer.CurrentY = iLine * oFontH
           Printer.Print StrToStr(.TextMatrix(i, 1), 8)
           '申請案號
           Printer.CurrentX = PLeft(2)
           Printer.CurrentY = iLine * oFontH
           Printer.Print StrToStr(.TextMatrix(i, 2), 8)
           '案件名稱
           Printer.CurrentX = PLeft(3)
           Printer.CurrentY = iLine * oFontH
           Printer.Print StrToStr(.TextMatrix(i, 3), 24)
           '申請日
           Printer.CurrentX = PLeft(4)
           Printer.CurrentY = iLine * oFontH
           Printer.Print StrToStr(.TextMatrix(i, 4), 5)
           iLine = iLine + 1
           oCount = oCount + 1
       Next i
       Printer.CurrentX = 0
       Printer.CurrentY = iLine * oFontH
       Printer.Print String(200, "-")
       Printer.CurrentX = 0
       Printer.CurrentY = (iLine + 1) * oFontH
       Printer.Print "共 " & oCount & " 筆"
   End With
   Printer.EndDoc
End Sub

Sub PrintData2()
Dim i As Integer
On Error GoTo 0

   If Printer.Orientation <> 1 Then
       Printer.Orientation = 1
   End If
   GetPleft
   SeekTemp(1) = ""
   SeekTemp(2) = ""
   PrintT ""
   PrintT2
   With GRD1
       For i = 1 To .Rows - 1
           strTemp(1) = StrToStr(.TextMatrix(i, 0), 4)
           strTemp(2) = StrToStr(.TextMatrix(i, 1), 4)
           strTemp(3) = StrToStr(.TextMatrix(i, 2), 5)
           strTemp(4) = StrToStr(.TextMatrix(i, 3), 25)
           strTemp(5) = StrToStr(.TextMatrix(i, 4), 4)
           strTemp(6) = StrToStr(.TextMatrix(i, 5), 4)
           strTemp(7) = StrToStr(.TextMatrix(i, 6), 4)
           If SeekTemp(1) <> .TextMatrix(i, 0) Then
               SeekTemp(1) = .TextMatrix(i, 0)
           Else
               strTemp(1) = ""
           End If
           If SeekTemp(2) <> .TextMatrix(i, 1) Then
               SeekTemp(2) = .TextMatrix(i, 1)
           Else
               strTemp(2) = ""
           End If
           If iLine > MaxLine Then
               Printer.NewPage
               strTemp(1) = SeekTemp(1)
               strTemp(2) = SeekTemp(2)
               PrintT ""
               PrintT2
           End If
           '業務區
           Printer.CurrentX = PLeft(1)
           Printer.CurrentY = iLine * oFontH
           Printer.Print strTemp(1)
           '智權人員
           Printer.CurrentX = PLeft(2)
           Printer.CurrentY = iLine * oFontH
           Printer.Print strTemp(2)
           '客戶編號
           Printer.CurrentX = PLeft(3)
           Printer.CurrentY = iLine * oFontH
           Printer.Print strTemp(3)
           '客戶名稱
           Printer.CurrentX = PLeft(4)
           Printer.CurrentY = iLine * oFontH
           Printer.Print strTemp(4)
           '已重新委任件數
           Printer.CurrentX = PLeft(5) + 900 - Printer.TextWidth(strTemp(5))
           Printer.CurrentY = iLine * oFontH
           Printer.Print strTemp(5)
           '待文件件數
           Printer.CurrentX = PLeft(6) + 900 - Printer.TextWidth(strTemp(6))
           Printer.CurrentY = iLine * oFontH
           Printer.Print strTemp(6)
           '不必重新委任件數
           Printer.CurrentX = PLeft(7) + 900 - Printer.TextWidth(strTemp(7))
           Printer.CurrentY = iLine * oFontH
           Printer.Print strTemp(7)
           iLine = iLine + 1
       Next i
   End With
   Printer.EndDoc
End Sub

Sub GetPleft()
   If txt1(6) = "1" Then
       PLeft(1) = 0
       PLeft(2) = 1750
       PLeft(3) = 4200
       PLeft(4) = 10000
   Else
       PLeft(1) = 0
       PLeft(2) = 1000
       PLeft(3) = 2000
       PLeft(4) = 3300
       PLeft(5) = 8000
       PLeft(6) = 9100
       PLeft(7) = 10200
   End If
End Sub

Sub PrintT(oStr As String)
Dim oStrByCU As String
Dim rsTmp As New ADODB.Recordset
Dim strSql  As String

   Printer.Font.Size = 22
   Printer.CurrentX = Printer.ScaleWidth / 2 - Printer.TextWidth("重新委任案件") / 2
   Printer.CurrentY = 600
   Printer.Print "重新委任案件 "
   Printer.Font.Size = 12
   iLine = 4
   Printer.CurrentX = 9000
   Printer.CurrentY = iLine * oFontH
   Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   iLine = 5
   If Trim(oStr) <> "" Then
       If rsTmp.State = 1 Then rsTmp.Close
       strSql = "select NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) from customer where cu01||cu02='" & oStr & "' "
       rsTmp.CursorLocation = adUseClient
       rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If rsTmp.RecordCount > 0 Then
           oStrByCU = CheckStr(rsTmp.Fields(0))
       Else
           oStrByCU = ""
       End If
       rsTmp.Close
       Printer.CurrentX = 0
       Printer.CurrentY = iLine * oFontH
       Printer.Print "客戶：" & oStr & "  " & oStrByCU
       iLine = 7
   End If
End Sub

Sub PrintT1(oStr As String, oStr2 As String)
MyStart:
   Printer.CurrentX = 0
   Printer.CurrentY = iLine * oFontH
   Printer.Print IIf(oStr = "10", "已重新委任案件", IIf(oStr = "20", "待處裡案件", "不必重新委任案件"))
   iLine = iLine + 1
   If iLine > MaxLine Then
       Printer.NewPage
       oPage = oPage + 1
       PrintT oStr2
       GoTo MyStart
   End If
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "本所案號"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "申請案號"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "案件名稱"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "申請日"
   iLine = iLine + 1
   If iLine > MaxLine Then
       Printer.NewPage
       oPage = oPage + 1
       PrintT oStr2
       GoTo MyStart
   End If
   Printer.CurrentX = 0
   Printer.CurrentY = iLine * oFontH
   Printer.Print String(200, "=")
   iLine = iLine + 1
   If iLine > MaxLine Then
       Printer.NewPage
       oPage = oPage + 1
       PrintT oStr2
       GoTo MyStart
   End If
End Sub

Sub PrintT2()
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * oFontH
   Printer.Print ""
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print ""
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * oFontH
   Printer.Print ""
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * oFontH
   Printer.Print ""
   Printer.CurrentX = PLeft(5) + ((PLeft(6) - PLeft(5)) / 2) - (Printer.TextWidth("已重新") / 2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "已重新"
   Printer.CurrentX = PLeft(6) + ((PLeft(7) - PLeft(6)) / 2) - (Printer.TextWidth("待文件") / 2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "待文件"
   Printer.CurrentX = PLeft(7) + ((11300 - PLeft(7)) / 2) - (Printer.TextWidth("不必重新") / 2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "不必重新"
   iLine = iLine + 1
   Printer.CurrentX = PLeft(1)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "業務區"
   Printer.CurrentX = PLeft(2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "智權人員"
   Printer.CurrentX = PLeft(3)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "客戶編號"
   Printer.CurrentX = PLeft(4)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "客戶名稱"
   Printer.CurrentX = PLeft(5) + ((PLeft(6) - PLeft(5)) / 2) - (Printer.TextWidth("委任件數") / 2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "委任件數"
   Printer.CurrentX = PLeft(6) + ((PLeft(7) - PLeft(6)) / 2) - (Printer.TextWidth("件數") / 2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "件數"
   Printer.CurrentX = PLeft(7) + ((11300 - PLeft(7)) / 2) - (Printer.TextWidth("委任件數") / 2)
   Printer.CurrentY = iLine * oFontH
   Printer.Print "委任件數"
   iLine = iLine + 1
   Printer.CurrentX = 0
   Printer.CurrentY = iLine * oFontH
   Printer.Print String(200, "=")
   iLine = iLine + 1
End Sub
