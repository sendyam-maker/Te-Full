VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040324_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "實審請求期限屆滿前通知函"
   ClientHeight    =   3480
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7980
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7980
   Begin VB.TextBox text1 
      Height          =   264
      Index           =   2
      Left            =   1560
      MaxLength       =   1
      TabIndex        =   0
      Top             =   2760
      Width           =   375
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   5520
      TabIndex        =   1
      Top             =   30
      Width           =   972
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   6525
      TabIndex        =   2
      Top             =   30
      Width           =   1215
   End
   Begin MSForms.ComboBox cbo 
      Height          =   300
      Index           =   1
      Left            =   1320
      TabIndex        =   4
      Top             =   1725
      Width           =   6465
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11404;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ComboBox cbo 
      Height          =   300
      Index           =   0
      Left            =   1320
      TabIndex        =   3
      Top             =   1320
      Width           =   6465
      VariousPropertyBits=   679479323
      DisplayStyle    =   7
      Size            =   "11404;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "法定期限："
      Height          =   180
      Index           =   5
      Left            =   4470
      TabIndex        =   17
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label Label1 
      Height          =   180
      Index           =   4
      Left            =   5420
      TabIndex        =   16
      Top             =   2460
      Width           =   1995
   End
   Begin VB.Label lblPA11 
      Height          =   180
      Left            =   1320
      TabIndex        =   15
      Top             =   585
      Width           =   2775
   End
   Begin VB.Label Label1 
      Height          =   180
      Index           =   3
      Left            =   1320
      TabIndex        =   14
      Top             =   2460
      Width           =   1995
   End
   Begin VB.Label Label1 
      Height          =   180
      Index           =   0
      Left            =   1320
      TabIndex        =   13
      Top             =   960
      Width           =   2235
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   12
      Top             =   960
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請案號："
      Height          =   180
      Index           =   0
      Left            =   330
      TabIndex        =   11
      Top             =   585
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "是否修改定稿：             (Y : Word )"
      Height          =   180
      Index           =   8
      Left            =   240
      TabIndex        =   10
      Top             =   2820
      Width           =   2670
   End
   Begin VB.Label Label1 
      Height          =   180
      Index           =   2
      Left            =   1320
      TabIndex        =   9
      Top             =   2130
      Width           =   1755
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "本所期限："
      Height          =   180
      Index           =   6
      Left            =   330
      TabIndex        =   8
      Top             =   2460
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利種類："
      Height          =   180
      Index           =   4
      Left            =   330
      TabIndex        =   7
      Top             =   2130
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "申請人１："
      Height          =   180
      Index           =   3
      Left            =   330
      TabIndex        =   6
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "專利名稱："
      Height          =   180
      Index           =   2
      Left            =   330
      TabIndex        =   5
      Top             =   1380
      Width           =   900
   End
End
Attribute VB_Name = "frm040324_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/20 改成Form2.0 (cbo)
'Create by Lydia 2015/07/20 實審請求期限屆滿前通知函
Option Explicit

Public m_NowPA11 As String '傳申請號
Public m_NowNo As String '傳案號

Dim m_ET01 As String '定稿別
Dim m_ET02 As String '總收文號(或本所案號&案件性質)
Dim m_ET03 As String '處理狀況
Dim m_strPA01 As String '本所案號
Dim m_strPA02 As String '本所案號
Dim m_strPA03 As String '本所案號
Dim m_strPA04 As String '本所案號
Dim m_strPA08 As String '專利種類
Dim m_strPA09 As String '申請國家
Dim m_strPA26 As String '申請人1
Dim m_strPA46 As String 'PCT案
Dim m_strNP08 As String '所限
Dim m_strNP09 As String '法限
Dim strSql As String
Dim m_NP01 As String '實審收文號
Dim m_LD18 As String '期限進度收文號
'Added by Lydia 2015/08/12
Dim m_bolFMP As Boolean '是否FMP案
Public m_bolRead As Boolean '是否有資料
Dim m_strPA75 As String 'Added by Morgan 2016/6/16
Dim m_bolFMP2 As Boolean 'Added by Lydia 2023/05/17 是否為寰華案
'Add By Sindy 2016/10/7
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
'2016/10/7 END
Dim stCP13 As String, stCP12 As String
Dim m_strPA57 As String 'Added by Lydia 2023/05/26

'Modified by Lydia 2015/08/12
'Private Sub cmdExit_Click()
Public Sub cmdExit_Click()
    Unload Me
    frm040324.Show
End Sub

Private Sub cmdok_Click()
Dim bolChk As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   If TxtValidate = False Then Exit Sub
   
   'Add By Sindy 2022/7/1
   'Mark by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail:可取消外專系統收件區，key來函承辦人掛程序人員，則按確定，信件會再打開一次的設定。
   'If m_strIR01 <> "" And Left(Pub_StrUserSt03, 2) = "F2" Then
   '   If PUB_ChkFileOpening2(Forms(0).Tmpfrm04010519.m_strFullFileName, "後續才能一併歸卷！") = True Then
   '      Exit Sub
   '   End If
   'End If
   ''2022/7/1 END
   'end 2023/05/17
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   
   If FormSave = False Then
      MsgBox "存檔失敗，請洽系統管理員 !", vbCritical
      Me.Enabled = True
      Screen.MousePointer = vbDefault
   Else
      
      If text1(2) = "Y" Then
         bolChk = True
      Else
         bolChk = False
      End If
      
      '列印定稿 -用一般來函
      m_ET01 = "07"
      'PCT案
      If m_strPA46 = "Y" Then
         m_ET03 = "01"
      Else
         m_ET03 = "00"
      End If

       m_ET02 = m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04 & "&1406"
      StartLetter m_ET01, m_ET02, m_ET03
      'Added by Lydia 2015/08/15 FMP案只印1份
      If m_bolFMP Then
         'Modified by Morgan 2025/4/11 FMP不再印紙本--品薇
         NowPrint m_ET02, m_ET01, m_ET03, bolChk, strUserNum, 0, , , , 1, , , , , , , , m_LD18, , , , , True
      Else
         NowPrint m_ET02, m_ET01, m_ET03, bolChk, strUserNum, 0, , , , , , , , , , , , m_LD18
      End If
      'end 2015/08/12
      
      'Added by Morgan 2016/6/17
      If Left(Pub_StrUserSt03, 1) <> "F" Then
         If bolChk Then
            frm1105_1.m_RecNo = m_LD18
            frm1105_1.m_PdfName = PUB_CaseNo2FileName(m_strPA01, m_strPA02, m_strPA03, m_strPA04) & ".1406.CUS.PDF"
            frm1105_1.Show
         End If
      End If
      'end 2016/6/17
       
      frm040324.Show
      frm040324.Clear
      Me.Enabled = True
      Screen.MousePointer = vbDefault
      'Add By Sindy 2017/12/29
      If Me.m_strIR01 <> "" Then
         Unload frm040324
         'Modify By Sindy 2022/5/20
         'frm04010519.GoNext
         Forms(0).Tmpfrm04010519.GoNext
         Set Forms(0).Tmpfrm04010519 = Nothing
         '2022/5/20 END
      End If
      '2017/12/29 END
      Unload Me
   End If
End Sub

Private Sub StartLetter(ByVal ET01 As String, ByVal ET02 As String, ByVal ET03 As String)
   Dim strTxt(1 To 10) As String
   Dim i As Integer, j As Integer
   Dim strFee As String
   
   EndLetter ET01, ET02, ET03, strUserNum
   i = 0
   
    strFee = PUB_GetYF0607(m_strPA09, m_strPA08, m_strPA26, "416", "1", "1", "1")
   
    i = i + 1
    strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','本所期限','" & m_strNP08 & "')"
    i = i + 1
    strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','法定期限','" & m_strNP09 & "')"
    i = i + 1
    strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','下一程序','416')"
    i = i + 1
    strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
                   "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','費用'," & strFee & ")"
    i = i + 1
    strExc(0) = Pub_Get416Period(m_strPA08, m_strPA09)
    strTxt(i) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) VALUES " & _
      "('" & ET01 & "','" & ET02 & "','" & ET03 & "','" & strUserNum & "','提實審期限','" & strExc(0) & "')"
    
    If Not ClsLawExecSQL(i, strTxt) Then
        MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
    End If

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   'Modified by Lydia 2015/08/12
   'ReadData
   m_bolRead = ReadData
   
   'Add By Sindy 2017/12/29
   m_strIR01 = frm040324.m_strIR01
   m_strIR02 = frm040324.m_strIR02
   m_strIR03 = frm040324.m_strIR03
   m_strIR04 = frm040324.m_strIR04
   If m_strIR01 <> "" Then
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   '2017/12/29 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   PUB_SendMailCache 'Added by Lydia 2023/05/17
   Set frm040324_2 = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
   TextInverse Me.text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
   Case 2 '是否修改定稿
      KeyAscii = UpperCase(KeyAscii)
      If KeyAscii <> 89 And KeyAscii <> 8 Then
         MsgBox "是否修改定稿只能輸入 Y !!!", vbExclamation + vbOKOnly
         KeyAscii = 0
      End If
   End Select
End Sub

Private Function ReadData() As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

    ReadData = False
    'Modified by Lydia 2015/08/12 判斷是否為FMP案
    'StrSQLa = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA09,PA08,PA11,PA26,PA46,PA57,CU04,CU05,CU06,CU88,CU89,CU90,PTM03,NP01,NP08,NP09 " & _
              "FROM PATENT,CUSTOMER,PATENTTRADEMARKMAP,NEXTPROGRESS WHERE " & ChgPatent(Replace(m_NowNo, "-", "")) & " AND PA11='" & m_NowPA11 & "' AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
              "AND '1'=PTM01(+) AND PA08=PTM02(+) AND NP02(+)=PA01 AND NP03(+)=PA02 AND NP04(+)=PA03 AND NP05(+)=PA04 AND NP07='" & 實體審查 & "'"
    'Modified by Morgan 2016/6/16 +pa75
    StrSQLa = "SELECT PA01,PA02,PA03,PA04,PA05,PA06,PA07,PA09,PA08,PA11,PA26,PA46,PA57,CU04,CU05,CU06,CU88,CU89,CU90,PTM03,NP01,NP08,NP09,CP12,PA75 " & _
              "FROM PATENT,CUSTOMER,PATENTTRADEMARKMAP,NEXTPROGRESS,CASEPROGRESS WHERE " & ChgPatent(Replace(m_NowNo, "-", "")) & " AND PA11='" & m_NowPA11 & "' AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
              "AND '1'=PTM01(+) AND PA08=PTM02(+) AND NP02(+)=PA01 AND NP03(+)=PA02 AND NP04(+)=PA03 AND NP05(+)=PA04 AND NP07(+)='" & 實體審查 & "'" & _
              "AND NP01=CP09(+)"
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        m_strPA01 = rsA.Fields("PA01").Value
        m_strPA02 = rsA.Fields("PA02").Value
        m_strPA03 = rsA.Fields("PA03").Value
        m_strPA04 = rsA.Fields("PA04").Value
        m_strPA08 = "" & rsA("PA08").Value
        m_strPA09 = "" & rsA("PA09").Value
        m_strPA26 = "" & rsA("PA26").Value
        m_strPA46 = "" & rsA("PA46").Value
        m_NP01 = "" & rsA("NP01").Value
        m_strNP08 = "" & rsA("NP08").Value
        m_strNP09 = "" & rsA("NP09").Value
        m_strPA75 = "" & rsA("PA75").Value 'Added by Morgan 2016/6/16
        m_strPA57 = "" & rsA("PA57").Value 'Added by Lydia 2023/05/26
        
        Me.Label1(0).Caption = m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04
        '申請號
        Me.lblPA11.Caption = rsA.Fields("PA11").Value
        
'Removed by Morgan 2016/2/1 來函不需控制閉卷不通知--韻丞
'        '是否閉卷
'        If "" & rsA.Fields("PA57") = "Y" Then
'           MsgBox "此案已閉卷!!!", vbExclamation + vbOKOnly
'           'Modified by Lydia 2015/08/12
'           'Unload Me
'           Exit Function
'        Else
'end 2016/2/1

        If m_strNP09 = "" Then
            MsgBox "本案無實體審查!!!", vbExclamation + vbOKOnly
           'Modified by Lydia 2015/08/12
           'Unload Me
           Exit Function
        End If
        '專利名稱(中-->英-->日)
        Me.cbo(0).AddItem "中：" & rsA.Fields("PA05").Value
        Me.cbo(0).AddItem "英：" & rsA.Fields("PA06").Value
        Me.cbo(0).AddItem "日：" & rsA.Fields("PA07").Value
        Me.cbo(0).ListIndex = 0
        '申請人1(中-->英-->日)
        Me.cbo(1).AddItem "中：" & rsA.Fields("CU04").Value
        Me.cbo(1).AddItem "英：" & Trim("" & rsA.Fields("CU05").Value & " " & rsA.Fields("CU88").Value & " " & rsA.Fields("CU89").Value & " " & rsA.Fields("CU90").Value & " ")
        Me.cbo(1).AddItem "日：" & rsA.Fields("CU06").Value
        Me.cbo(1).ListIndex = 0
        '專利種類

        Me.Label1(2).Caption = "" & rsA("PTM03").Value
        '所限,法限
        Me.Label1(3).Caption = ChangeTStringToTDateString(ChangeWStringToTString(m_strNP08))
        Me.Label1(4).Caption = ChangeTStringToTDateString(ChangeWStringToTString(m_strNP09))
   
        'Added by Lydia 2015/08/12
        'Modified by Morgan 2021/1/28
        'If m_strPA01 = "P" And rsA.Fields("PA09") <> "000" And Left("" & rsA.Fields("CP12"), 1) = "F" Then
        '從 Formsave 移來以便共用
        stCP13 = PUB_GetAKindSalesNo(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
        stCP12 = GetSalesArea(stCP13)
        If m_strPA01 = "P" And rsA.Fields("PA09") <> "000" And Left(stCP12, 1) = "F" Then
        'end 2021/1/28
           m_bolFMP = True
        Else
           m_bolFMP = False
        End If
         'Added by Lydia 2023/05/17 判斷寰華案
         m_bolFMP2 = False
         If m_bolFMP = True Then
            If PUB_FMPtoCheck(1, 2, Pub_strUserST05, m_strPA01, m_strPA02, m_strPA03, m_strPA04) = True Then
               m_bolFMP2 = True
            End If
         End If
         'end 2023/05/17
        ReadData = True
    Else
        MsgBox "基本檔無此案號資料!!!", vbExclamation + vbOKOnly
        'Modified by Lydia 2015/08/12
        'Unload Me
        Exit Function
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Function

Private Function TxtValidate() As Boolean
   TxtValidate = False
   '重複通知檢查
   If DupeCheck = False Then
      Exit Function
   End If
   TxtValidate = True
End Function

Private Function DupeCheck() As Boolean
   DupeCheck = True
   strExc(0) = "select cp10, cp27 from caseprogress where cp01='" & m_strPA01 & "' and cp02='" & m_strPA02 & "' and cp03='" & m_strPA03 & "'  and cp04='" & m_strPA04 & "'" & _
               " and cp10='1406' and cp57 is null order by cp27 desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If MsgBox("本案件已於 " & ChangeWStringToTDateString("" & RsTemp.Fields(1)) & " 通知請求期限屆滿，是否再次通知？", vbYesNo + vbDefaultButton2) = vbNo Then
            DupeCheck = False
      End If
   End If
End Function
Private Function FormSave() As Boolean
   
   Dim cp() As String
   ReDim cp(1 To TF_CP) As String
   Dim stUpdatePA As String
 
On Error GoTo ErrorHandler

   cnnConnection.BeginTrans
On Error GoTo ErrorHandler1
   
   cp(1) = m_strPA01
   cp(2) = m_strPA02
   cp(3) = m_strPA03
   cp(4) = m_strPA04
   cp(5) = strSrvDate(1)
   cp(9) = 主管機關來函
   cp(10) = "1406"
   cp(12) = stCP12
   cp(13) = stCP13
   cp(14) = strUserNum
   cp(27) = strSrvDate(1)
   cp(20) = "N"
   cp(26) = "N"
   cp(32) = "N"
   cp(43) = m_NP01
   cp(119) = strSrvDate(1)
   
   strSql = GetCPSQL(cp(), False)
   cnnConnection.Execute strSql, intI
   
   m_LD18 = cp(9)
   If m_strPA09 <> "000" Then
      '抓最新的AB類發文代理人更新
      Pub_UpdateFromMaxCP27 cp(1), cp(2), cp(3), cp(4)
      
      'Added by Morgan 2016/6/16
      If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And Left(Pub_StrUserSt03, 1) <> "F" Then
         PUB_AddLetterProgress m_LD18, 2, True, "", True, m_strPA26, "1406", m_strPA75
      End If
      'end 2016/6/16
   End If
   
   'Add by Sindy 2017/12/29
   If m_strIR01 <> "" Then
      'Modify By Sindy 2022/6/28 + , IIf(Pub_StrUserSt03 = "F22", m_LD18, "")
      'Modified by Lydia 2023/05/18 +不開啟附件, , , False
      PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "frm040324_2", IIf(Pub_StrUserSt03 = "F22", m_LD18, ""), , , False
   End If
   '2017/12/29 END
   'Added by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail
   'Modified by Lydia 2023/05/26 已閉卷不通知
   If m_bolFMP = True And m_bolFMP2 = True And m_strPA57 = "" Then
       'Modified by Lydia 2023/10/31 傳入C類收文號 m_LD18
       Call Pub_SetFMP2toCMail(m_strPA01, m_strPA02, m_strPA03, m_strPA04, cp(10), cp(14), m_LD18) '傳入相關收文的承辦人
       
   'Added by Morgan 2025/3/6 FMP非寰華也要通知--韻丞
   ElseIf m_bolFMP = True And m_bolFMP2 = False Then
      PUB_FMPCaseInform m_LD18, False, True, False
   'end 2025/3/6
   End If
   'end 2023/05/17

   cnnConnection.CommitTrans
   FormSave = True
   Exit Function
   
ErrorHandler1:
   cnnConnection.RollbackTrans
   
ErrorHandler:

End Function

