VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm12040129 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員客戶轉移作業"
   ClientHeight    =   3210
   ClientLeft      =   1935
   ClientTop       =   1605
   ClientWidth     =   5865
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3210
   ScaleWidth      =   5865
   Begin VB.TextBox txtCustNo 
      Height          =   264
      Index           =   1
      Left            =   2565
      MaxLength       =   9
      TabIndex        =   3
      Text            =   "X"
      Top             =   1800
      Width           =   972
   End
   Begin VB.TextBox txtCustNo 
      Height          =   264
      Index           =   0
      Left            =   1344
      MaxLength       =   9
      TabIndex        =   2
      Text            =   "X"
      Top             =   1800
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      Height          =   400
      Left            =   4812
      TabIndex        =   5
      Top             =   120
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   3984
      TabIndex        =   4
      Top             =   120
      Width           =   800
   End
   Begin VB.TextBox textNewNum 
      Height          =   264
      Left            =   1344
      TabIndex        =   1
      Top             =   1425
      Width           =   972
   End
   Begin VB.TextBox textOldNum 
      Height          =   264
      Left            =   1344
      TabIndex        =   0
      Top             =   1080
      Width           =   972
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "原智權人員未離職時，案件進度檔及承辦人工作進度資料不更新！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   270
      TabIndex        =   14
      Top             =   660
      Width           =   5220
   End
   Begin MSForms.TextBox textNewNum_2 
      Height          =   300
      Left            =   2424
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Width           =   3252
      VariousPropertyBits=   671105055
      Size            =   "5736;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textOldNum_2 
      Height          =   300
      Left            =   2424
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   1077
      Width           =   3252
      VariousPropertyBits=   671105055
      Size            =   "5736;529"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(所有客戶請勿改預設值)"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   3672
      TabIndex        =   11
      Top             =   1872
      Width           =   1920
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "    只更新本所期限未過期且未續辦之案件！"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   10
      Top             =   2760
      Width           =   3420
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "PS:2017/10/20發現下一程序資料有誤,執行前後都要檢查!!!!!!"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   240
      TabIndex        =   9
      Top             =   2520
      Width           =   4680
   End
   Begin VB.Line Line1 
      X1              =   2295
      X2              =   2610
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Label Label3 
      Caption         =   "客戶編號："
      Height          =   180
      Left            =   270
      TabIndex        =   8
      Top             =   1845
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "新智權人員："
      Height          =   180
      Left            =   270
      TabIndex        =   7
      Top             =   1470
      Width           =   1080
   End
   Begin VB.Label Label1 
      Caption         =   "原智權人員："
      Height          =   180
      Left            =   270
      TabIndex        =   6
      Top             =   1122
      Width           =   1080
   End
End
Attribute VB_Name = "frm12040129"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2021/12/10 Form2.0已修改(textOldNum_2,textNewNum_2)
'Memo By Sonia 2012/12/5 智權人員欄已修改
'2010/12/2 memo by sonia 員工編號欄已修改
'sonia 2010/9/15 日期欄已修改
Option Explicit

Private Sub Form_Load()
   textOldNum_2.BackColor = &H8000000F
   textNewNum_2.BackColor = &H8000000F
   
   MoveFormToCenter Me
End Sub

Private Sub cmdOK_Click()
   If CheckDataValid() = True Then
      'Add By Cheng 2002/05/23
      '重新檢查欄位有效性
      If TxtValidate = False Then Exit Sub
      ' 設定滑鼠游標為等待狀態
      Screen.MousePointer = vbHourglass
      ' 執行更新的作業
      '92.3.5 modify by sonia
      'OnSaveData
      If OnSaveData = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Screen.MousePointer = vbDefault: Exit Sub
      ' 設定滑鼠游標為預設
      Screen.MousePointer = vbDefault
      ' 清除欄位內容
      textOldNum = Empty
      textOldNum_2 = Empty
      textNewNum = Empty
      textNewNum_2 = Empty
      ' 設定輸入欄位
      textOldNum.SetFocus
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Function OnSaveData() As Boolean
Dim strDepart As String
Dim strSql As String
Dim n1Count As Integer
Dim n2Count As Integer
Dim n3Count As Integer
Dim n4Count As Integer
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim stCon As String 'Added by Morgan 2015/3/31
Dim bolLeave As Boolean 'Add by Amy 2022/08/24
Dim stNewNo As String 'Add by Amy 2022/10/26
   
   '92.5.23 Add By sonia
   OnSaveData = False
   On Error GoTo ErrorHandler
   cnnConnection.BeginTrans
   
   ' 取得新智權人員的業務區別
   '2010/10/21 modify by sonia 改抓st15(98024)
   'strDepart = GetStaffDepartment(textNewNum)
   strDepart = GetST15(textNewNum)
   n1Count = 0: n2Count = 0: n3Count = 0: n4Count = 0
      
   'Added by Morgan 2015/3/31
   If txtCustNo(0) <> "X" And txtCustNo(0) <> "" Then
      stCon = stCon & " and cu01>='" & txtCustNo(0) & "'"
   End If
   
   If txtCustNo(1) <> "X" And txtCustNo(1) <> "" Then
      stCon = stCon & " and cu01<='" & txtCustNo(1) & "ZZZ'"
   End If
   
   'Add by Amy 2022/08/24 舊智權人員是否已離職
   bolLeave = ChkStaffST04(Me.textOldNum.Text, False)
   
   'Modify By Cheng 2003/09/29
   '先抓舊智權人員的客戶
   StrSqlB = "Select * From Customer Where CU13='" & Me.textOldNum.Text & "' " & stCon
   rsB.CursorLocation = adUseClient
   rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   While Not rsB.EOF
       'Modify by Amy 2022/10/26 +新智權為S部門且客戶狀態是解散/廢止/撤銷/死亡,更新為區無效
       If Left(strDepart, 1) = "S" And ("" & rsB.Fields("cu80") = "解散" Or "" & rsB.Fields("cu80") = "廢止" Or "" & rsB.Fields("cu80") = "撤銷" Or "" & rsB.Fields("cu80") = "死亡") Then
            stNewNo = GetAreaEmpNo(strDepart)
       Else
            stNewNo = Me.textNewNum.Text
       End If
'      'Add by Amy 2022/08/24 舊智權人員(非F部門)離職,客戶未發文進度及歷程同步修改
'      If bolLeave = True Then
'          Call Pub_ChangeSaleUpdCP13("" & rsB.Fields("cu01"), Me.textOldNum.Text, Me.textNewNum.Text)
'      End If
      '更新某一客戶的下一程序智權人員資料
      StrSQLa = "Select NP01, NP07, NP22 From NextProgress, Patent Where NP02=PA01 And NP03=PA02 And NP04=PA03 And NP05=PA04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And PA26='" & rsB("CU01").Value & rsB("CU02").Value & "' "
      StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22 From NextProgress, Trademark Where NP02=TM01 And NP03=TM02 And NP04=TM03 And NP05=TM04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And TM23='" & rsB("CU01").Value & rsB("CU02").Value & "' "
      StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22 From NextProgress, Lawcase Where NP02=LC01 And NP03=LC02 And NP04=LC03 And NP05=LC04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And LC11='" & rsB("CU01").Value & rsB("CU02").Value & "' "
      StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22 From NextProgress, Hirecase Where NP02=HC01 And NP03=HC02 And NP04=HC03 And NP05=HC04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And HC05='" & rsB("CU01").Value & rsB("CU02").Value & "' "
      StrSQLa = StrSQLa & " Union Select NP01, NP07, NP22 From NextProgress, ServicePractice Where NP02=SP01 And NP03=SP02 And NP04=SP03 And NP05=SP04 And NP06 Is Null And NP08>=" & strSrvDate(1) & " And SP08='" & rsB("CU01").Value & rsB("CU02").Value & "' "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      While Not rsA.EOF
          'edit by nickc 2005/10/28 程序的案子不用更新
          'StrSQLa = "Update Nextprogress Set NP10='" & Me.textNewNum.Text & "' Where NP01='" & rsA.Fields(0).Value & "' And NP07='" & rsA.Fields(1).Value & "' And NP22=" & rsA.Fields(2).Value
          '2010/3/23 MODIFY BY SONIA 程序管制案件性質不印改以strNpSqlOfNoSalesDuty控制
          'Modify by Amy 2022/10/26 原:Me.textNewNum.Text
          StrSQLa = "Update Nextprogress Set NP10='" & stNewNo & "' Where NP01='" & rsA.Fields(0).Value & "' And NP07='" & rsA.Fields(1).Value & "' And NP22=" & rsA.Fields(2).Value & strNpSqlOfNoSalesDuty
          
          cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2008/12/5 +控制來函期限通知的 Trigger 不被觸發
          cnnConnection.Execute StrSQLa, n2Count
          cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2008/12/5 +控制來函期限通知的 Trigger 不被觸發
          n2Count = n2Count + n2Count
          rsA.MoveNext
      Wend
      'Add by Amy 2022/10/26
      If Left(strDepart, 1) = "S" Then
            '更新客戶基本資料檔
            strSql = "UPDATE CUSTOMER SET CU12 = '" & strDepart & "', CU13 = '" & stNewNo & "' " & _
                        ",CU79=DECODE(CU79,NULL,'" & strSrvDate(2) & "改智權人員,原為'||CU13,'" & strSrvDate(2) & "改智權人員,原為'||CU13||';'||CU79) " & _
                        "WHERE CU13 = '" & textOldNum & "' And cu01='" & rsB.Fields("cu01") & "' And cu02='" & rsB.Fields("cu02") & "' "
            cnnConnection.Execute strSql, n1Count
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
      rsB.MoveNext
   Wend
   If rsB.State <> adStateClosed Then rsB.Close
   'Modify by Amy 2022/10/26 非智權部整批更新
   If Left(strDepart, 1) <> "S" Then
        ' 更新客戶基本資料檔
        strSql = "UPDATE CUSTOMER SET CU12 = '" & strDepart & "', CU13 = '" & textNewNum & "' " & _
                 ",CU79=DECODE(CU79,NULL,'" & strSrvDate(2) & "改智權人員,原為'||CU13,'" & strSrvDate(2) & "改智權人員,原為'||CU13||';'||CU79) WHERE CU13 = '" & textOldNum & "' " & stCon
        cnnConnection.Execute strSql, n1Count
   End If
      
   'Modify By Sindy 2022/9/20 更新進度檔和EP的資料:
   '秀玲反應會慢，且發生跑葉招進85026客戶轉至中四備用20042，大約5~6分鐘。
   '而且有個問題，現在進度檔還有20筆葉招進的未發文資料，這些客戶在今天以前已經轉給其他智權人員。
   '考慮這種情形，還是改寫，先抓所有未發文未閉卷未銷卷進度，依案件現在之客戶檔的智權人員去更新進度檔的智權人員。
   If bolLeave = True Then
      StrSqlB = "Select cu01||cu02 CUID,cu13,st02,st04 From Customer,staff where cu01||cu02 in(" & _
       "SELECT TM23 FROM TRADEMARK WHERE (TM01,TM02,TM03,TM04) in (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Me.textOldNum.Text & "' AND CP158=0 AND CP159=0) AND TM30||TM57 is null" & _
       " union SELECT PA26 FROM PATENT WHERE (PA01,PA02,PA03,PA04) in (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Me.textOldNum.Text & "' AND CP158=0 AND CP159=0) AND PA58||PA108 is null" & _
       " union SELECT SP08 FROM SERVICEPRACTICE WHERE (SP01,SP02,SP03,SP04) in (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Me.textOldNum.Text & "' AND CP158=0 AND CP159=0) AND SP16||SP61 is null" & _
       " union SELECT LC11 FROM LAWCASE WHERE (LC01,LC02,LC03,LC04) in (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Me.textOldNum.Text & "' AND CP158=0 AND CP159=0) AND LC09||LC34 is null" & _
       " union SELECT HC05 FROM HIRECASE WHERE (HC01,HC02,HC03,HC04) in (SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP13='" & Me.textOldNum.Text & "' AND CP158=0 AND CP159=0) AND HC10||HC19 is null" & _
       ") and cu13 is not null and cu13=st01(+) and st04='1'"
      rsB.CursorLocation = adUseClient
      rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
      While Not rsB.EOF
         Call Pub_ChangeSaleUpdCP13(rsB.Fields("CUID"), Me.textOldNum.Text, rsB.Fields("cu13"))
         rsB.MoveNext
      Wend
      If rsB.State <> adStateClosed Then rsB.Close
   End If
   '2022/9/20 END
  
    '92.10.30 還原 by sonia
    If stCon = "" Then 'Added by Morgan 2015/3/31
        ' 更新下一程序檔
        'modify by sonia 2021/11/17 加入And NP08>=" & strSrvDate(1) & "條件
        strSql = "UPDATE NEXTPROGRESS SET NP10 = '" & textNewNum & "' " & _
                "WHERE NP10 = '" & textOldNum & "' And NP08>=" & strSrvDate(1) & " AND " & _
                        "(NP06 IS NULL OR NP06 = '' OR NP06 = ' ')"
           
        cnnConnection.Execute "begin user_data.user_notrigger:=1; end;" 'Add by Morgan 2008/12/5 +控制來函期限通知的 Trigger 不被觸發
        cnnConnection.Execute strSql, n3Count
        cnnConnection.Execute "begin user_data.user_notrigger:=0; end;" 'Add by Morgan 2008/12/5 +控制來函期限通知的 Trigger 不被觸發
           
    End If 'Added by Morgan 2015/3/31
   
    'add by sonia 2020/5/5 更新智權文件寄送確認未確認資料
    'modify by sonia 2022/5/10 +LP15<>'Y'條件
    strSql = "UPDATE LETTERPROGRESS SET LP06 = '" & textNewNum & "' WHERE LP06 = '" & textOldNum & "' AND NVL(LP07,0)=0 AND LP15<>'Y' "
    cnnConnection.Execute strSql, n4Count
    'end 2020/5/5
   
   If n1Count + n2Count + n3Count + n4Count <= 0 Then
      MsgBox "無此智權人員的相關資料可更新", vbOKOnly + vbInformation, "智權人員客戶轉移作業"
   End If
   
OnSaveData = True
'92.5.23 Add By sonia
cnnConnection.CommitTrans
   
   'Modify By Sindy 2022/9/19 Mark
'   'Add By Sindy 2016/5/20
'   '智權人員離職時,調整待會稿區正在送會中及會圖中的收受者
'   Call PUB_SalseLeaveUpEEP05(textOldNum, , False)
   Set rsB = Nothing
Exit Function

ErrorHandler:
    cnnConnection.RollbackTrans
    'Add by Morgan 2008/12/5 因為 Rollback 不會還原 package 的變數設定所以要人工執行還原的語法
    cnnConnection.Execute "begin user_data.user_notrigger:=0; end;"
    OnSaveData = False

End Function

' 檢查輸入的資料是否完整
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   CheckDataValid = False

   ' 原智權人員編號
   If IsEmptyText(textOldNum) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入原智權人員編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldNum.SetFocus
      GoTo EXITSUB
   End If
   
   ' 新智權人員編號
   If IsEmptyText(textNewNum) = True Then
      strTit = "檢核資料"
      strMsg = "請輸入新智權人員編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNewNum.SetFocus
      GoTo EXITSUB
   End If
   
   ' 原智權人員代號不可為新智權人員代號
   If textOldNum = textNewNum Then
      strTit = "檢核資料"
      strMsg = "新智權人員編號不可為原智權人員編號"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNewNum.SetFocus
      GoTo EXITSUB
   End If
   
   ' 原智權人員代號不存在
   textOldNum_2 = GetStaffName(textOldNum, True)
   If IsEmptyText(textOldNum_2) = True Then
      strTit = "檢核資料"
      strMsg = "原智權人員代號不存在或已離職"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textOldNum.SetFocus
      GoTo EXITSUB
   End If
   
   ' 新智權人員代號不存在
   textNewNum_2 = GetStaffName(textNewNum, False)
   If IsEmptyText(textNewNum_2) = True Then
      strTit = "檢核資料"
      strMsg = "新智權人員代號不存在或已離職"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      textNewNum.SetFocus
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Private Sub Form_Unload(Cancel As Integer)
   'Add By Cheng 2002/07/18
   Set frm12040129 = Nothing
End Sub

Private Sub textNewNum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textOldNum_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 原智權人員代號
Private Sub textOldNum_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textOldNum_2 = Empty
   If IsEmptyText(textOldNum) = False Then
      textOldNum_2 = GetStaffName(textOldNum, True)
      If IsEmptyText(textOldNum_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "原智權人員代號不存在或已離職"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textOldNum_GotFocus
      End If
   End If
End Sub

' 新智權人員代號
Private Sub textNewNum_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   textNewNum_2 = Empty
   If IsEmptyText(textNewNum) = False Then
      textNewNum_2 = GetStaffName(textNewNum, False)
      If IsEmptyText(textNewNum_2) = True Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "新智權人員代號不存在或已離職"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textNewNum_GotFocus
      End If
   End If
End Sub

Private Sub textOldNum_GotFocus()
   InverseTextBox textOldNum
End Sub

Private Sub textNewNum_GotFocus()
   InverseTextBox textNewNum
End Sub

'Add By Cheng 2002/05/23
Private Function TxtValidate() As Boolean
Dim objTxt As Object
Dim ii As Integer
Dim Cancel As Boolean

TxtValidate = False
If Me.textNewNum.Enabled = True Then
   Cancel = False
   textNewNum_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

If Me.textOldNum.Enabled = True Then
   Cancel = False
   textOldNum_Validate Cancel
   If Cancel = True Then
      Exit Function
   End If
End If

TxtValidate = True
End Function

Private Sub txtCustNo_GotFocus(Index As Integer)
   txtCustNo(Index).SelStart = 2
End Sub

Private Sub txtCustNo_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
