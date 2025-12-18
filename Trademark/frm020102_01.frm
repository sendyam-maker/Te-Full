VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm020102_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "發文"
   ClientHeight    =   5730
   ClientLeft      =   2080
   ClientTop       =   1540
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   9310
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   4032
      Left            =   96
      TabIndex        =   14
      Top             =   1656
      Width           =   9132
      _ExtentX        =   16104
      _ExtentY        =   7108
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體-ExtB"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.OptionButton radio 
      Caption         =   "申請案號 :"
      Height          =   252
      Index           =   2
      Left            =   120
      TabIndex        =   13
      Top             =   1260
      Width           =   1332
   End
   Begin VB.TextBox textTM12 
      Height          =   264
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   12
      Top             =   1260
      Width           =   2892
   End
   Begin VB.CommandButton cmdExtent 
      Caption         =   "延期(&D)"
      Height          =   400
      Left            =   5568
      TabIndex        =   10
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "發文資料(&F)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6396
      TabIndex        =   11
      Top             =   70
      Width           =   1200
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Height          =   400
      Left            =   7620
      TabIndex        =   8
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   8448
      TabIndex        =   9
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox textCP09 
      Height          =   264
      Left            =   1560
      MaxLength       =   9
      TabIndex        =   2
      Top             =   660
      Width           =   2892
   End
   Begin VB.TextBox textTM02_2 
      Height          =   264
      Left            =   3000
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Visible         =   0   'False
      Width           =   372
   End
   Begin VB.OptionButton radio 
      Caption         =   "收文號 :"
      Height          =   252
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   660
      Width           =   1332
   End
   Begin VB.OptionButton radio 
      Caption         =   "本所案號 :"
      Height          =   252
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1332
   End
   Begin VB.TextBox textTM01 
      Height          =   264
      Left            =   1560
      MaxLength       =   3
      TabIndex        =   3
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox textTM03 
      Height          =   264
      Left            =   3360
      MaxLength       =   1
      TabIndex        =   6
      Top             =   960
      Width           =   372
   End
   Begin VB.TextBox textTM04 
      Height          =   264
      Left            =   3720
      MaxLength       =   2
      TabIndex        =   7
      Top             =   960
      Width           =   732
   End
   Begin VB.TextBox textTM02 
      Height          =   264
      Left            =   2280
      MaxLength       =   6
      TabIndex        =   4
      Top             =   960
      Width           =   1092
   End
End
Attribute VB_Name = "frm020102_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2023/10/13 Form2.0已修改 ; grdList從MSFlexGrid改為MSHFlexGrid
'Memo by Amy 2021/12/21 Form2.0已修改 grdList
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/16 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

' 使用者所選取的查詢方式是收文號還是本所案號
Dim m_KeySel As Integer
' 使用者所選取的收文號
Dim m_CP09 As String
' 使用者所選取的列其位置
Dim m_CurrSel As Integer
' 本所案號
Dim m_TM01 As String
Dim m_TM02 As String
Dim m_TM03 As String
Dim m_TM04 As String
Dim m_CP10 As String '案件性質
Dim m_intofirst As Boolean
Dim m_TM10 As String 'Add by Amy 2014/10/16 申請國家
Dim m_TM16 As String 'Added by Lydia 2016/01/19 已准駁
Public bolIsEMPFlow As Boolean 'Add By Sindy 2018/5/3 是否為電子承辦簽核


Public Sub Clear()
   textCP09 = Empty
   textTM12 = Empty 'Added by Morgan 2023/1/3
   'textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
   InitialGrdList
   radio(0).Value = True
   radio(1).Value = False
   radio_Click 0
End Sub

Public Sub Clear1()
   textCP09 = Empty
   textTM12 = Empty 'Added by Morgan 2023/1/3
   'textTM01 = Empty
   textTM02 = Empty
   textTM02_2 = Empty
   textTM03 = Empty
   textTM04 = Empty
   InitialGrdList
   cmdQuery.Default = True
   'Add By Sindy 2009/06/04
   'Modified by Morgan 2023/6/30
   'radio(1).Value = True
   'radio_Click 1
   'Call textTM02.SetFocus
   textTM01 = Empty
   radio(2).Value = True
   radio_Click 2
   textTM12.SetFocus
   'end 2023/6/30
   '2009/06/04 End
End Sub

' 90.07.23 add
' 檢查原案的本所及法定期限是否存在
Private Function CheckCP0607(ByVal bShowMsg As Boolean) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckCP0607 = True
   
   strSql = "SELECT CP06,CP07 FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CP06")) = True Then
         CheckCP0607 = False
      Else
         If IsEmptyText(rsTmp.Fields("CP06")) = True Then
            CheckCP0607 = False
         Else
            If rsTmp.Fields("CP06") = "0" Then
               CheckCP0607 = False
            End If
         End If
      End If
      
      If IsNull(rsTmp.Fields("CP07")) = True Then
         CheckCP0607 = False
      Else
         If IsEmptyText(rsTmp.Fields("CP07")) = True Then
            CheckCP0607 = False
         Else
            If rsTmp.Fields("CP07") = "0" Then
               CheckCP0607 = False
            End If
         End If
      End If
   End If
   rsTmp.Close
   
   If CheckCP0607 = False And bShowMsg = True Then
      strTit = "檢核資料"
      strMsg = "原案之案件進度資料無本所及法定期限, 無法執行延期作業!"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   End If
   Set rsTmp = Nothing
End Function

'延期按鈕
Private Sub cmdExtent_Click()
Dim frmNext As Form
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   If CheckDataValid = True Then
      'Add By Cheng 2002/07/15
      '所點選的案件性質不可為"延期"
      If PUB_CPKindDelay(Me.grdList.TextMatrix(Me.grdList.row, 6), "T") Then
         Exit Sub
      End If
      'Add By Cheng 2002/07/12 若案件已閉卷, 不可發文
      If PUB_CaseClosedCP09(Me.grdList.TextMatrix(Me.grdList.row, 6)) = True Then
         Exit Sub
      End If
      
      '2006/3/20 ADD BY SONIA 若專用期間已過期但發文案件性質非延展,補正時, 不可發文
      'modify by sonia 2019/5/28 +剔除延期303(FCT-043529的補正延期)
      If Me.grdList.TextMatrix(Me.grdList.row, 7) <> "102" And Me.grdList.TextMatrix(Me.grdList.row, 7) <> "201" And Me.grdList.TextMatrix(Me.grdList.row, 7) <> "303" Then
         StrSQLa = "SELECT TM22 FROM TRADEMARK,CASEPROGRESS WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "'"
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            'edit by nickc 2006/06/26 半年內皆可以補繳
            'If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ServerDate Then
                'MsgBox "此案件專用期間已過, 不可執行發文作業!!!", vbExclamation + vbOKOnly
            If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(ServerDate))) Then
               MsgBox "此案件專用期間已過半年, 不可執行發文作業!!!", vbExclamation + vbOKOnly
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               Exit Sub
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      'add by sonia 2019/5/28 延期303可能為延展補正延期(FCT-043529),若已過專用期則提醒不必限制
      ElseIf Me.grdList.TextMatrix(Me.grdList.row, 8) = "303" Then
         StrSQLa = "SELECT TM22 FROM TRADEMARK,CASEPROGRESS WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "'"
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ServerDate Then
               MsgBox "此案件專用期間已過, 請確認是否仍要發文延期 !!!"
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      'end 2019/5/28
      End If
      '2006/3/20 END
      
      'Add By Sindy 2018/5/3
      '檢查是否有承辦歷程是否有產生承辦單可以發文
      If PUB_IsEmpFlowIsSend(m_CP09) = False Then
         Exit Sub
      End If
      '2018/5/3 END
      
      If CheckCP0607(True) = True Then
         'Add By Cheng 2002/01/31
         frm020102_11.bln_From_Frm020102_01_BtnExt = True
         Set frmNext = frm020102_11
         ' 顯示下一個畫面
         'If IsObject(frmNext) = True Then
            frmNext.SetData 0, m_CP09, True
            Me.Hide
            frmNext.Show
            frmNext.QueryData
         'End If
            StrSQLa = "Select CP10 From Caseprogress Where CP09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                m_CP10 = "" & rsA.Fields(0).Value
            Else
                m_CP10 = ""
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            '顯示商標基本資料的畫面
            '若案件性質為異議, 評定, 廢止時, 不要檢查是否要補基本檔資料
            'Modified by Lydia 2023/10/13 601+627, 603+629, 605+623
            'If m_CP10 <> "601" And m_CP10 <> "603" And m_CP10 <> "605" And m_CP10 <> "308" Then
            If m_CP10 <> "601" And m_CP10 <> "627" And m_CP10 <> "603" And m_CP10 <> "629" And m_CP10 <> "605" And m_CP10 <> "623" And m_CP10 <> "308" Then
                ShowMaintainForm m_CP09
            End If
      End If
   End If
End Sub

'確定按鈕
Private Sub cmdok_Click()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
    
   'Added by Lydia 2017/07/20 提示無點選
   If "" & Me.grdList.TextMatrix(Me.grdList.row, 6) = "" Or Me.grdList.TextMatrix(Me.grdList.row, 6) = "收文號" Then
      MsgBox "請點選進度!", vbCritical
      Exit Sub
   End If
   'end 2017/07/20
   
   'add by sonia 2019/4/17 開放外商程序可發文TF中間程序(即子案)
   If Left(Pub_StrUserSt03, 2) = "F1" Then
      StrSQLa = "select cp09,cp01,cp02,cp03,cp04 from caseprogress where cp09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "' and cp01='TF' and cp03<>'0' and cp04<>'00'"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount = 0 Then
         MsgBox "外商程序只可發文TF案之中間程序 !!!", vbCritical
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         Exit Sub
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   'end 2019/4/17
   
   'add by sonia 2015/9/15 無商標圖時提醒,台灣T案不限案件性質,非台灣案僅申請,分割才要提醒
   'modify by sonia 2016/5/13 除FCT外所有案件無商標圖時都提醒(桂英)
   If Me.grdList.TextMatrix(Me.grdList.row, 9) = "FCT" Then
   'modify by sonia 2016/5/13 除FCT外所有案件無商標圖時都提醒(桂英)
   'ElseIf m_TM10 = 台灣國家代號 Or Me.grdList.TextMatrix(Me.grdList.row, 7) = "101" Or Me.grdList.TextMatrix(Me.grdList.row, 7) = "308" Then
   Else
      'modify by sonia 2019/4/17 TF子案都抓母案TF-000780-1-07
      'StrSQLa = "select cp09,ibf01 from caseprogress,ImgByteFile where cp09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "' and cp01=ibf01(+) and cp02=ibf02(+) and cp03=ibf03(+) and cp04=ibf04(+)"
      If Me.grdList.TextMatrix(Me.grdList.row, 9) = "TF" Then
         StrSQLa = "select cp09,ibf01 from caseprogress,ImgByteFile where cp09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "' and cp01=ibf01(+) and substr(cp02,1,5)||'0'=ibf02(+) and '0'=ibf03(+) and '00'=ibf04(+)"
      Else
         StrSQLa = "select cp09,ibf01 from caseprogress,ImgByteFile where cp09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "' and cp01=ibf01(+) and cp02=ibf02(+) and cp03=ibf03(+) and cp04=ibf04(+)"
      End If
      'end 2019/4/17
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 And IsNull(rsA.Fields(1).Value) Then
         MsgBox "此案件沒有代表圖 !!!", vbInformation
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If
   'end 2015/9/15
   
   ' 檢查是否資料已完全輸入
   If CheckDataValid = True Then
      'Added by Lydia 2016/01/19 台灣案的註冊費發文時,判斷案件為已准
      If (Me.grdList.TextMatrix(Me.grdList.row, 9) = "FCT" Or Me.grdList.TextMatrix(Me.grdList.row, 9) = "T") And m_TM10 = 台灣國家代號 And m_TM16 <> "1" And Me.grdList.TextMatrix(Me.grdList.row, 7) = "717" Then
         MsgBox "本案尚未核准不可繳註冊費!", vbCritical
         Exit Sub
      End If
      'end 2016/01/19
      
      'Added by Lydia 2015/11/24 管控台灣延展案102,系統日不得早於"延展期滿前6個月"的第一天
      If m_TM10 = 台灣國家代號 And Me.grdList.TextMatrix(Me.grdList.row, 7) = "102" And Not IsNull(Me.grdList.TextMatrix(Me.grdList.row, 10)) Then
         'Modified by Lydia 2017/06/01 延展期滿日期改用模組控制; 因為下午可發次日,所以多判斷前一工作天
         'If strSrvDate(1) < CompWorkDay(2, CompDate(1, -6, Me.grdList.TextMatrix(Me.grdList.row, 10)), 1) Then
         If strSrvDate(1) < CompWorkDay(2, PUB_Get102DeadLine("3", Me.grdList.TextMatrix(Me.grdList.row, 10)), 1) Then
            MsgBox "台灣延展案發文時,系統日不得早於延展期滿前6個月的第一天!", vbCritical
            Exit Sub
         End If
      End If
      'end 2015/11/24
        
      'Add by Amy 2014/10/16 +T大陸案分割控制
      If Me.grdList.TextMatrix(Me.grdList.row, 9) = "T" And m_TM10 = 大陸國家代號 _
         And Me.grdList.TextMatrix(Me.grdList.row, 7) = "308" And Me.grdList.TextMatrix(Me.grdList.row, 8) = "Y" Then
        MsgBox "T大陸分割新案不可發文，請由母案分割程序發文!", vbExclamation
        Exit Sub
      End If
      'end 2014/10/16
      
      'add by sonia 2018/11/19
      '異議案逾法定期限不可發文
      'Modified by Lydia 2023/10/13 +627
      If (Me.grdList.TextMatrix(Me.grdList.row, 7) = "601" Or Me.grdList.TextMatrix(Me.grdList.row, 7) = "627") And Not IsNull(Me.grdList.TextMatrix(Me.grdList.row, 10)) Then
         If strSrvDate(1) > Me.grdList.TextMatrix(Me.grdList.row, 10) Then
            MsgBox "異議案逾法定期限不可發文!", vbCritical
            Exit Sub
         End If
      End If
      '廢止案不可提早發文,管制期限存在CP46,發文存檔時要清除
      'modify by sonia 2018/12/3 +623部分廢止
      If (Me.grdList.TextMatrix(Me.grdList.row, 7) = "605" Or Me.grdList.TextMatrix(Me.grdList.row, 7) = "623") And Not IsNull(Me.grdList.TextMatrix(Me.grdList.row, 10)) Then
         If strSrvDate(1) < Me.grdList.TextMatrix(Me.grdList.row, 10) Then
            MsgBox "廢止案未達管制期限 (公告" & IIf(m_TM10 = "000", "滿", "期滿加") & "三年 " & ChangeTStringToTDateString(ChangeWStringToTString(Me.grdList.TextMatrix(Me.grdList.row, 10))) & ") 不可提早發文!", vbCritical
            Exit Sub
         End If
      End If
      'end 2018/11/19
        
      'add by sonia 2017/3/22 因為下一程序催審無人員,故控制沒有承辦人不可發文
      If Me.grdList.TextMatrix(Me.grdList.row, 3) = "" Then
         MsgBox "尚未分案, 沒有承辦人, 不可發文！", vbExclamation
         Exit Sub
      End If
      'end 2017/3/22
      'Add By Sindy 2023/3/29
      If Left(Me.grdList.TextMatrix(Me.grdList.row, 6), 1) = "A" Then
         strSql = "select cp09,cp140,cp157 from caseprogress where cp09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If Val("" & RsTemp.Fields("cp157")) = 0 And "" & RsTemp.Fields("cp140") <> "" Then
               MsgBox "尚未分案, 沒有北所分案日期, 不可發文！", vbExclamation
               Exit Sub
            End If
         End If
      End If
      '2023/3/29 END
      
      'Add By Cheng 2002/07/12
      '若案件已閉卷, 不可發文
      'modify by sonia 2025/9/11 +1728收款寄證 T-254847
      If Me.grdList.TextMatrix(Me.grdList.row, 7) <> "725" Then   '2012/5/3 ADD BY SONIA 退費不控制 T-158649
         If PUB_CaseClosedCP09(Me.grdList.TextMatrix(Me.grdList.row, 6)) = True Then
            Exit Sub
         End If
      End If '2012/5/3 ADD BY SONIA
      
      '2006/3/20 ADD BY SONIA 若專用期間已過期但發文案件性質非延展,補正時, 不可發文
      If Me.grdList.TextMatrix(Me.grdList.row, 7) <> "102" And Me.grdList.TextMatrix(Me.grdList.row, 7) <> "201" Then
         StrSQLa = "SELECT TM22 FROM TRADEMARK,CASEPROGRESS WHERE CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & Me.grdList.TextMatrix(Me.grdList.row, 6) & "'"
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            'edit by nickc 2006/06/26 半年內皆可以補繳
            'If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ServerDate Then
                'MsgBox "此案件專用期間已過, 不可執行發文作業!!!", vbExclamation + vbOKOnly
            If "" & rsA.Fields(0).Value <> "" And rsA.Fields(0).Value < ChangeWDateStringToWString(DateAdd("m", -6, ChangeWStringToWDateString(ServerDate))) Then
               MsgBox "此案件專用期間已過半年, 不可執行發文作業!!!", vbExclamation + vbOKOnly
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               Exit Sub
            End If
         End If
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
      End If
      '2006/3/20 END
      
      'Modify By Sindy 2024/1/23 改為共用函數
      If PUB_ChkCP141IsSend(m_CP09) = False Then
         Exit Sub
      End If
'      '91.7.16此段錯誤, 正確的控制在DisplayNextForm的ShowMaintainForm m_CP09
'      ' 檢查是否要顯示商標基本檔資料維護的畫面
'      'If CheckJumpFrm020501() = True Then
'      '   DisplayFrm020501
'      'Else
'      ' 檢查是否已收款
'      'Added by Lydia 2015/11/24
'      'Modify By Sindy 2023/4/21 +,cp142
'      'Modify By Sindy 2023/12/11 +,cp164
'      strExc(0) = "select cp06,nvl(cp79,0) cp79,cp141,cp142,cp164 from caseprogress where cp09='" & m_CP09 & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If "" & RsTemp.Fields("cp141") = "2" And RsTemp.Fields("cp79") > 0 Then
'             If PUB_ChkPaidByCP09(m_CP09) = False Then   'Added by Morgan 2016/8/23 出納繳款確認後就可送件
'                If IsNull(RsTemp.Fields("cp06")) Or "" & RsTemp.Fields("cp06") > strSrvDate(1) Then
'                   MsgBox "此案智權人員欲管控收款後才可送件，暫不可發文！"
'                   Exit Sub
'                End If
'             End If
'         ElseIf "" & RsTemp.Fields("cp141") = "3" Then
'            'Modify By Sindy 2023/12/11
'            '1=當天
'            If "" & RsTemp.Fields("cp164") = "1" And "" & RsTemp.Fields("cp142") <> "" And "" & RsTemp.Fields("cp142") > strSrvDate(1) Then
'               MsgBox "本案需於指定日" & ChangeWStringToTDateString(RsTemp.Fields("cp142")) & "方可發文！"
'               Exit Sub
'            '3=之後
'            ElseIf "" & RsTemp.Fields("cp164") = "3" And "" & RsTemp.Fields("cp142") <> "" And "" & RsTemp.Fields("cp142") >= strSrvDate(1) Then
'               MsgBox "本案需於指定日" & ChangeWStringToTDateString(RsTemp.Fields("cp142")) & "之後方可發文！"
'               Exit Sub
'            'Add By Sindy 2023/4/21
'            ElseIf "" & RsTemp.Fields("cp142") <> strSrvDate(1) Then
'               If MsgBox("本案已設定指定送件日為" & ChangeWStringToTDateString(RsTemp.Fields("cp142")) & "，但該日期與系統日不符，是否仍要發文？", vbYesNo + vbDefaultButton2) = vbNo Then
'                  Exit Sub
'               End If
'            '2023/4/21 END
'            End If
'         End If
'      End If
'      'end 2015/11/24
      '2024/1/23 END
      
      'Add By Sindy 2018/5/3
      '檢查是否有承辦歷程是否有產生承辦單可以發文
      If PUB_IsEmpFlowIsSend(m_CP09) = False Then
         Exit Sub
      End If
      '2018/5/3 END
      
      'Added by Morgan 2022/12/21
      '電子商標註冊證，承辦人需由承辦歷程送核判主管判發。反之，領取紙本商標註冊證則無需經核判。--嘉雯
      strExc(1) = Me.grdList.TextMatrix(Me.grdList.row, 7)
      If m_TM10 = "000" And m_CP09 < "C" And InStr(TMCertPtyList, strExc(1)) > 0 Then
         strExc(0) = "select cp01,cp02,cp03,cp04,cp10 from caseprogress where cp09='" & m_CP09 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If PUB_TWCertPty(RsTemp("cp01"), RsTemp("cp10"), RsTemp("cp02"), RsTemp("cp03"), RsTemp("cp04")) = True Then
               If PUB_GetCertType(RsTemp("cp01"), RsTemp("cp02"), RsTemp("cp03"), RsTemp("cp04")) = "1" Then
                  If PUB_ChkEmpFlowExists(m_CP09, EMP_送件) = False Then
                     MsgBox "本案設定為電子商標註冊證，【" & Me.grdList.TextMatrix(Me.grdList.row, 2) & "】需由承辦歷程送核判主管判！", vbExclamation
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
      'end 2022/12/21
      
         'Modified by Morgan 2016/8/23 出納繳款確認後就可送件
         'If CheckIfFinishCP79() = False Then
'Remove by Lydia 2018/08/22  (應收帳款管控)取消預定收款日,改成付款週期=>不發email
'         If CheckIfFinishCP79() = False And PUB_ChkPaidByCP09(m_CP09) = False Then
'         'end 2016/8/23
'            ' 若未收款時則顯示 frm030101_02 的畫面
'            DisplayFrm020102_02
'         Else
            'Move by Lydia 2016/01/04 凡收款後才可送件,若未送件不可發文(移到上方)
            
            ' 若已收款時則依案件性質顯示下一個畫面
            DisplayNextForm
'         End If
      'End If
'end  2018/08/22
   End If
End Sub

Private Sub Form_Activate()
   'Add By Sindy 2009/06/04
   If m_intofirst = True Then
      'Modified by Morgan 2023/6/29
      'radio(1).Value = True
      'radio_Click 1
      radio(2).Value = True
      radio_Click 2
      'end 2023/6/29
      m_intofirst = False
   End If
   '2009/06/04 End
   
   'Add By Sindy 2023/2/16 ex:T-182332([其他].CUS.來自承辦歷程送入)
   If m_CP09 <> "" Then PUB_UpdateLP03 m_CP09
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   Initial
   UpdateCtrlState
   InitialGrdList
   m_intofirst = True
End Sub

Private Sub Initial()
   ' 預設由申請案號來取得資料
   'Modify By Sindy 2009/06/04
   'm_KeySel = 0
   m_KeySel = 1
End Sub

' 按下結束離開按紐
Private Sub cmdExit_Click()
    'Add By Cheng 2004/04/08
'move to unload by nick 2004/10/22
'    '列印接洽接案單
''    PUB_PrintCaseCloseSheet strUserNum
'    PUB_PrintCaseCloseSheet strUserNum, "0", False, False
'    '刪除暫存資料
'    PUB_DeleteCaseCloseSheet strUserNum
    Unload Me
End Sub
' 按下查詢按紐
Private Sub cmdQuery_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   ' 先檢查該輸入的資料是否有全部輸入
   Select Case m_KeySel
      ' 依收文號
      Case 0:
         If IsEmptyText(textCP09) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入收文號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
      ' 依本所案號
      Case 1:
         If IsEmptyText(textTM01) = True Or IsEmptyText(textTM02) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入本所案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
         If textTM01 = "TF" Then
            If IsEmptyText(textTM03) = True And IsEmptyText(textTM04) = True Then
               frm020102_03.SetData 0, textTM01, True
               frm020102_03.SetData 1, textTM02 & textTM02_2, False
               frm020102_03.Show
               frm020102_03.QueryData
               Me.Hide
               GoTo EXITSUB
            End If
         End If
         
      'Added by Morgan 2023/1/3
      ' 依申請案號
      Case 2:
         If IsEmptyText(textTM12) = True Then
            strTit = "資料檢核"
            strMsg = "請輸入申請案號"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            GoTo EXITSUB
         End If
   End Select
   ' 查詢資料
   If QueryData = False Then
      'strTit = "資料查詢"
      'strMsg = "沒有符合條件的資料"
      'nResponse = MsgBox(strMsg, vbOKOnly, strTit)
   Else
      cmdok.Default = True
   End If
EXITSUB:
End Sub

' 檢查是否已收款
Private Function CheckIfFinishCP79() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   CheckIfFinishCP79 = True
    'Add By Cheng 2002/10/29
    '若申請國家不是台灣者, 不需檢查是否已收款
    strSql = "Select TM10 From CaseProgress, TradeMark Where CP01=TM01(+) AND CP02=TM02(+) AND CP03=TM03(+) AND CP04=TM04(+) AND CP09='" & m_CP09 & "'  AND TM10<>'000' "
    'Add By Cheng 2003/02/10
    '服務業務資料的搜尋
    strSql = strSql & " union Select SP09 From CaseProgress, ServicePractice Where CP01=SP01(+) AND CP02=SP02(+) AND CP03=SP03(+) AND CP04=SP04(+) AND CP09='" & m_CP09 & "'  AND SP09<>'000' "
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    Set rsTmp = Nothing
    rsTmp.CursorLocation = adUseClient
    rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    If rsTmp.RecordCount <= 0 Then
        If rsTmp.State <> adStateClosed Then rsTmp.Close
        Set rsTmp = Nothing
        Exit Function
    End If
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    Set rsTmp = Nothing
       
   ' 查詢案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CP79")) = False Then
         If IsEmptyText(rsTmp.Fields("CP79")) = False Then
            If rsTmp.Fields("CP79") <> "0" Then
               CheckIfFinishCP79 = False
            End If
         End If
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function

'Mark by Amy 2021/12/21 和Morgan確認不使用了
'' 顯示向智權人員發EMail的畫面
'Private Sub DisplayFrm020102_02()
'   frm020102_02.SetData 0, m_CP09, True
'   Me.Hide
'   frm020102_02.Show
'   frm020102_02.QueryData
'End Sub

Public Sub RefreshData()
   Dim bQuery As Boolean
   bQuery = QueryData
End Sub

' 查詢資料庫
Public Function QueryData() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Dim strSql As String
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   Dim rsTmp As New ADODB.Recordset
   
   QueryData = False
   m_CP09 = Empty
   InitialGrdList
   
   ' 組成SQL語法
   Select Case m_KeySel
      ' 依收文號
      Case 0:
         ' 檢查案件進度檔, 系統類別必須為FCT, 且必須為未輸入發文日, 且未輸入取消收文日期的A,B類收文號
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP09 = '" & textCP09 & "' AND " & _
                        "(CP01 LIKE 'T%' OR " & _
                        "CP01 = 'FCT')"
                        
      ' 依本所案號
      Case 1:
         strCP01 = Trim(textTM01)
         strCP02 = Trim(textTM02)
         If textTM01 = "TF" Then: strCP02 = strCP02 & textTM02_2
         strCP03 = Trim(textTM03)
         If IsEmptyText(strCP03) = True Then: strCP03 = "0"
         strCP04 = Trim(textTM04)
         If IsEmptyText(strCP04) = True Then: strCP04 = "00"
         strSql = "SELECT * FROM CaseProgress " & _
                  "WHERE CP01 = '" & strCP01 & "' AND " & _
                        "CP02 = '" & strCP02 & "' AND " & _
                        "CP03 = '" & strCP03 & "' AND " & _
                        "CP04 = '" & strCP04 & "' "
                  
      'Added by Morgan 2023/1/3
      ' 依申請案號
      Case 2:
         strSql = "SELECT C.* FROM Trademark,CaseProgress C " & _
                  "WHERE TM12='" & textTM12 & "' AND cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and " & _
                        "(TM01 LIKE 'T%' OR TM01 = 'FCT')"
   End Select
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 列出所有資料
   If rsTmp.RecordCount > 0 Then
      If ListData(rsTmp) = True Then
         'cmdOK.SetFocus
         QueryData = True
      '92.7.4 ADD BY SONIA
      Else
         strTit = "資料查詢"
         strMsg = "沒有符合條件的資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      '92.7.4 END
      End If
   End If
   rsTmp.Close
   
EXITSUB:
   Set rsTmp = Nothing
End Function

' 列出所有符合條件的資料
Private Function ListData(ByRef rsTmp As ADODB.Recordset) As Boolean
   Dim nRow As Integer
      
   ListData = False
   m_TM10 = "": m_TM16 = ""  'Added by Lydia 2016/01/19
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   rsTmp.MoveFirst
   Do While rsTmp.EOF = False
      '收文號不為A,B類的不予計入
      Select Case Mid(rsTmp.Fields("CP09"), 1, 1)
         'Modified by Lydia 2016/12/22 +D類
         'Case "A", "B":
         Case "A", "B", "D":
         Case Else: GoTo NextRecord
      End Select
      '尚未輸入發文日
      If IsNull(rsTmp.Fields("CP27")) = False Then
         If IsEmptyText(rsTmp.Fields("CP27")) = False Then
            If rsTmp.Fields("CP27") <> "0" Then: GoTo NextRecord
         End If
      End If
      '尚未輸入取消收文日期
      If IsNull(rsTmp.Fields("CP57")) = False Then
         If IsEmptyText(rsTmp.Fields("CP57")) = False Then
            If rsTmp.Fields("CP57") <> "0" Then: GoTo NextRecord
         End If
      End If
      
      grdList.Rows = grdList.Rows + 1
      nRow = grdList.Rows - 1
      ' 收文日欄位
      If IsNull(rsTmp.Fields("CP05")) = False Then
         grdList.TextMatrix(nRow, 1) = TAIWANDATE(rsTmp.Fields("CP05"))
      End If
      ' 案件性質
      '910723
      If IsNull(rsTmp.Fields("CP10")) = False Then
         strExc(1) = rsTmp.Fields("CP01")
         strExc(2) = rsTmp.Fields("CP02")
         strExc(3) = rsTmp.Fields("CP03")
         strExc(4) = rsTmp.Fields("CP04")
         'edit by nickc 2007/02/06 不用 dll 了
         'If objPublicData.GetSystemKind(strExc(1), intI) Then
         If ClsPDGetSystemKind(strExc(1), intI) Then
            'Modified by Lydia 2016/01/19 +目前准駁TM16
            If intI = 2 Then
               strExc(0) = "SELECT TM10,TM16 FROM TRADEMARK WHERE TM01='" & strExc(1) & "' AND TM02='" & strExc(2) & "' AND TM03='" & strExc(3) & "' AND TM04='" & strExc(4) & "'"
            Else
               strExc(0) = "SELECT SP09,'' FROM SERVICEPRACTICE WHERE SP01='" & strExc(1) & "' AND SP02='" & strExc(2) & "' AND SP03='" & strExc(3) & "' AND SP04='" & strExc(4) & "'"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))   'edit by nickc 2007/02/06 不用 dll 了   = objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_TM10 = RsTemp.Fields(0) 'Add by Amy 2014/10/16
               m_TM16 = "" & RsTemp.Fields(1) 'Added by Lydia 2016/01/19
               If RsTemp.Fields(0) < "010" Then
                  grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 0)
               Else
                  grdList.TextMatrix(nRow, 2) = GetCaseTypeName(rsTmp.Fields("CP01"), rsTmp.Fields("CP10"), 1)
               End If
            End If
         End If
      End If
      ' 承辦人
      If IsNull(rsTmp.Fields("CP14")) = False Then
         'Modified by Lydia 2018/04/03 人員離職也會顯示 +true
         grdList.TextMatrix(nRow, 3) = GetStaffName(rsTmp.Fields("CP14"), True)
      End If
      ' 智權人員
      If IsNull(rsTmp.Fields("CP13")) = False Then
         grdList.TextMatrix(nRow, 4) = GetStaffName(rsTmp.Fields("CP13"))
      End If
      ' 進度備註
      If IsNull(rsTmp.Fields("CP64")) = False Then
         grdList.TextMatrix(nRow, 5) = rsTmp.Fields("CP64")
      End If
      ' 收文號
      If IsNull(rsTmp.Fields("CP09")) = False Then
         grdList.TextMatrix(nRow, 6) = rsTmp.Fields("CP09")
      End If
      '2006/3/20 ADD BY SONIA
      ' 案件性質
      If IsNull(rsTmp.Fields("CP10")) = False Then
         grdList.TextMatrix(nRow, 7) = rsTmp.Fields("CP10")
      End If
      '2006/3/20 END
      'Add by Amy 2014/10/16 +CP31 CP01
      grdList.TextMatrix(nRow, 8) = "" & rsTmp.Fields("CP31")
      grdList.TextMatrix(nRow, 9) = "" & rsTmp.Fields("CP01")
      'end 2014/10/16
      'Add By Sindy 2010/12/27 判斷有相關總收文號才做
      ' 相關總收文號
      If IsNull(rsTmp.Fields("CP43")) = False Then
         '案件性質
         grdList.TextMatrix(nRow, 2) = grdList.TextMatrix(nRow, 2) & PUB_GetRelateCasePropertyName(grdList.TextMatrix(nRow, 6), "1")
      End If
      '2010/12/27 End
      'Added by Lydia 2015/11/24 +法定期限
      grdList.TextMatrix(nRow, 10) = "" & rsTmp.Fields("CP07")
      'add by sonia 2018/11/20 廢止案不可提早發文,管制期限存在CP46,發文存檔時要清除
      If IsNull(rsTmp.Fields("CP10")) = False Then
         'modify by sonia 2018/12/3 +623部分廢止
         If rsTmp.Fields("CP10") = "605" Or rsTmp.Fields("CP10") = "623" Then
            grdList.TextMatrix(nRow, 10) = "" & rsTmp.Fields("CP46")
         End If
      End If
      'end 2018/11/20
      ListData = True
NextRecord:
      rsTmp.MoveNext
   Loop
   'Added by Lydia 2023/10/13
   If grdList.Rows >= 2 Then
      grdList.FixedRows = 1
   End If
   'end 2023/10/13
      
   ' 顯示符合的所有資料
   grdList.Refresh
   ' 設定第一筆為被選取的狀態
   grdList_SetSelection 1
   
EXITSUB:
End Function

' 更新控制項的狀態
Private Sub UpdateCtrlState()
   Select Case m_KeySel
      Case 0:
         EnableTextBox textCP09, True
         EnableTextBox textTM12, False 'Added by Morgan 2023/1/3
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         textTM02_2.Visible = False
      Case 1:
         EnableTextBox textTM12, False 'Added by Morgan 2023/1/3
         EnableTextBox textCP09, False
         EnableTextBox textTM01, True
         EnableTextBox textTM02, True
         EnableTextBox textTM03, True
         EnableTextBox textTM04, True
         textTM01_Validate False
      
      'Added by Morgan 2023/1/3
      Case 2:
         EnableTextBox textTM12, True
         EnableTextBox textTM01, False
         EnableTextBox textTM02, False
         EnableTextBox textTM03, False
         EnableTextBox textTM04, False
         EnableTextBox textCP09, False
         textTM02_2.Visible = False
         
   End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
   bolIsEMPFlow = False 'Add By Sindy 2018/5/3
   '列印接洽接案單
'   PUB_PrintCaseCloseSheet strUserNum
   PUB_PrintCaseCloseSheet strUserNum, "0", False, False
   '刪除暫存資料
   PUB_DeleteCaseCloseSheet strUserNum
   
   'Add By Sindy 2023/2/16 ex:T-182332([其他].CUS.來自承辦歷程送入)
   If m_CP09 <> "" Then PUB_UpdateLP03 m_CP09
   
   'Add By Cheng 2002/07/18
   Set frm020102_01 = Nothing
End Sub

'Add By Cheng 2002/01/10
Private Sub grdList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If grdList.Rows > 1 Then
   If grdList.row > 0 Then
      m_CP09 = grdList.TextMatrix(grdList.row, 6)
      Me.cmdok.SetFocus
   End If
End If
grdList_ShowSelection
End Sub

' 使用者按下所選取的項目
Public Sub radio_Click(Index As Integer)
'********** 90.11.23 nick
   If frm020102_01.Visible = True Then
   m_KeySel = Index
   UpdateCtrlState
   ' 90.07.25 modify
   Select Case Index
      Case 0:
         'Modify By Sindy 2018/6/29 + textCP09.Visible = True And bolIsEMPFlow = False
         If textCP09.Visible = True And bolIsEMPFlow = False Then textCP09.SetFocus
      Case 1:
         'Modify By Sindy 2018/6/29 + textCP09.Visible = True And bolIsEMPFlow = False
         If textTM01.Visible = True And bolIsEMPFlow = False Then textTM01.SetFocus
      
      'Added by Morgan 2023/1/3
      Case 2:
         If textTM12.Visible = True And bolIsEMPFlow = False Then textTM12.SetFocus
   End Select
   End If
   '**********
End Sub

Private Sub textCP09_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   '2005/8/29 ADD BY SONIA
   cmdQuery.Default = True
End Sub

Private Sub textTM01_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

' 檢查系統別
Private Sub textTM01_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   Cancel = False
   textTM02.MaxLength = 6
   If IsEmptyText(textTM01) = False Then
      If Mid(textTM01, 1, 1) = "T" Then
         Select Case textTM01
            Case "TF":
               textTM02_2.Visible = True
               textTM02_2.Locked = False
               textTM02_2.TabStop = True
               textTM02.MaxLength = 5
            Case Else:
               textTM02_2.Visible = False
               textTM02_2.Locked = True
               textTM02_2.TabStop = False
               textTM02.MaxLength = 6
         End Select
      Else
         Select Case textTM01
            Case "FCT":
               textTM02_2.Visible = False
               textTM02_2.Locked = True
               textTM02_2.TabStop = False
               textTM02.MaxLength = 6
            Case Else
               Cancel = True
               strTit = "資料檢核"
               strMsg = "本所案號中的系統別不正確"
               nResponse = MsgBox(strMsg, vbOKOnly, strTit)
               textTM01_GotFocus
         End Select
      End If
   Else
      textTM02_2.Visible = False
      textTM02_2.Locked = True
      textTM02_2.TabStop = False
      textTM02.MaxLength = 6
   End If
End Sub

' 初始化 GridList
Private Sub InitialGrdList()
   Dim nIndex As Integer
   grdList.Clear
   grdList.Rows = 1
   'Modified by Lydia 2015/11/24 +11
   'grdList.Cols = 10 'Modify by Amy 2014/10/16 原8
   grdList.Cols = 11
   grdList.ColWidth(0) = 300
   grdList.row = 0
   grdList.col = 1
   grdList.Text = "收文日"
   grdList.ColWidth(1) = 1200
   grdList.col = 2
   grdList.Text = "案件性質"
   grdList.ColWidth(2) = 1000
   grdList.col = 3
   grdList.Text = "承辦人"
   grdList.ColWidth(3) = 1200
   grdList.col = 4
   grdList.Text = "智權人員"
   grdList.ColWidth(4) = 1000
   grdList.col = 5
   grdList.Text = "進度備註"
   grdList.ColWidth(5) = 1200
   ' 收文號欄位 (隱藏欄位)
   grdList.col = 6
   grdList.Text = "收文號"
   grdList.ColWidth(6) = 0
   '2006/3/20 ADD BY SONIA
   grdList.col = 7
   grdList.Text = "CP10"
   grdList.ColWidth(7) = 0
   'Add by Amy 2014/10/16 T大陸案分割控制
   grdList.col = 8 '是否新案件
   grdList.Text = "CP31"
   grdList.ColWidth(8) = 0
   grdList.col = 9
   grdList.Text = "CP01"
   grdList.ColWidth(9) = 0
   'end 2014/10/16
   'Added by Lydia 2015/11/24 +法定期限
   grdList.col = 10
   grdList.Text = "CP07"
   grdList.ColWidth(10) = 0
End Sub

' 設定Grid List的一列為選取的狀態
Private Sub grdList_SetSelection(ByVal nSel As Integer)
   If nSel > 0 And nSel < grdList.Rows And grdList.Rows >= 2 Then
      grdList.row = nSel
      grdList_SelChange
   End If
End Sub

Private Sub grdList_SelChange()
   If grdList.Rows > 1 Then
      If grdList.row > 0 Then
         m_CP09 = grdList.TextMatrix(grdList.row, 6)
         Me.cmdok.SetFocus
      End If
   End If
   grdList_ShowSelection
   'cmdOK.SetFocus
End Sub

' 將GridList所選取的列反白, 並將未選取的列設成一般顏色
Private Sub grdList_ShowSelection()
   Dim nCurrSel As Integer
   Dim nCol As Integer
   
   nCurrSel = grdList.row
   
   ' 與前一選擇的列位置相同則不處理
   If m_CurrSel = grdList.row Then
      Dim nOldCol As Integer
      nOldCol = grdList.col
      grdList.col = 1
      If grdList.CellBackColor <> &H8000000D Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H8000000D Then grdList.CellBackColor = &H8000000D
            If grdList.CellForeColor <> &H80000005 Then grdList.CellForeColor = &H80000005
         Next nCol
      End If
      grdList.col = nOldCol
      GoTo EXITSUB
   End If
   
   ' 將原先選取的列回復到正常的顏色
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      If grdList.CellBackColor <> &H80000005 Then
         For nCol = 1 To grdList.Cols - 1
            grdList.col = nCol
            If grdList.CellBackColor <> &H80000005 Then: grdList.CellBackColor = &H80000005
            If grdList.CellForeColor <> &H80000008 Then: grdList.CellForeColor = &H80000008
         Next nCol
      End If
      grdList.col = 0
   End If
   ' 設定成所選取的列
   m_CurrSel = nCurrSel
   ' 將所選取的列反白
   If m_CurrSel > 0 And m_CurrSel < grdList.Rows Then
      grdList.row = m_CurrSel
      grdList.col = 1
      For nCol = 1 To grdList.Cols - 1
         grdList.col = nCol
         grdList.CellBackColor = &H8000000D
         grdList.CellForeColor = &H80000005
      Next nCol
      cmdok.SetFocus
      grdList.col = 0
   End If
EXITSUB:
End Sub

' 顯示下一個畫面
Public Sub DisplayNextForm()
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strCP10 As String
   Dim strCP01 As String, strCP31 As String
   Dim frmNext As Form
   Dim bNext As Boolean
   
   bNext = False
   strCP10 = Empty
   strCP31 = Empty 'Add By Sindy 2011/7/14
   ' 組成SQL語法
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 列出所有資料
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("CP01")) = False Then
         strCP01 = rsTmp.Fields("CP01")
      End If
      If IsNull(rsTmp.Fields("CP10")) = False Then
         strCP10 = rsTmp.Fields("CP10")
      End If
      'Add By Sindy 2011/7/14
      If IsNull(rsTmp.Fields("CP31")) = False Then
         strCP31 = rsTmp.Fields("CP31")
      End If
   End If
   rsTmp.Close

'CANCEL BY SONIA 2014/7/18 後面都有寫,此處無用
'   ' 先預設非商標基本檔的畫面為其它服務業務
'   If strCP01 <> "T" And strCP01 <> "TF" And strCP01 <> "CFT" And strCP01 <> "FCT" Then
'      Select Case strCP10
'         ' 變更  2010/10/14 加302
'         Case "301", "302":
'            Set frmNext = frm020102_07
'            bNext = True
'         Case Else:
'            Set frmNext = frm020102_21
'            bNext = True
'      End Select
'   End If
'END 2014/7/18
   
   Select Case strCP01
      Case "TR":
         Select Case strCP10
            ' 變更  2010/10/14 加302
            Case "301", "302":
               Set frmNext = frm020102_07
               bNext = True
            ' 商業司查詢
            Case "808":
               Set frmNext = frm020102_05
               bNext = True
            ' 其它
            Case Else:
               Set frmNext = frm020102_21
               bNext = True
         End Select
      Case "TM":
         Select Case strCP10
            ' 變更  2010/10/14 加302
            Case "301", "302":
               Set frmNext = frm020102_07
               bNext = True
            '2011/6/9 ADD BY SONIA
            Case "602":
               Set frmNext = frm020102_16
               bNext = True
            ' 其它
            Case Else:
               Set frmNext = frm020102_18
               bNext = True
         End Select
      Case "TB":
         Select Case strCP10
            ' 變更  2010/10/14 加302
            Case "301", "302":
               Set frmNext = frm020102_07
               bNext = True
            'Add By Cheng 2002/06/14
            ' 轉讓
            Case "501":
               Set frmNext = frm020102_08
               bNext = True
            ' 其它
            Case Else:
               Set frmNext = frm020102_19
               bNext = True
         End Select
      Case "TC":
         Select Case strCP10
            ' 變更  2010/10/14 加302
            Case "301", "302":
               Set frmNext = frm020102_07
               bNext = True
            '2014/7/18 ADD BY SONIA TC-010635
            ' 轉讓
            Case "501":
               Set frmNext = frm020102_08
               bNext = True
            'END 2014/7/18
            ' 其它
            Case Else:
               Set frmNext = frm020102_20
               bNext = True
         End Select
      Case "TD", "TT":
         Select Case strCP10
            ' 變更  2010/10/14 加302
            Case "301", "302":
               Set frmNext = frm020102_07
               bNext = True
            '2011/6/9 ADD BY SONIA
            Case "602":
               Set frmNext = frm020102_16
               bNext = True
            ' 其它
            Case Else:
               Set frmNext = frm020102_21
               bNext = True
         End Select
      Case "TS":
         Select Case strCP10
            ' 查名
            Case "001":
               Set frmNext = frm020102_05
               bNext = True
            ' 變更  2010/10/14 加302
            Case "301", "302":
               Set frmNext = frm020102_07
               bNext = True
            ' 其它
            Case Else:
               Set frmNext = frm020102_21
               bNext = True
         End Select
      Case Else
         Select Case strCP10
            'Modify By Cheng 2003/02/21
            '加申請中文證明
'            ' 查名, 申請, 延展, 補換發證書, 申請英文證明, 刊登廣告
'            Case "101", "102", "103", "304", "702":
            ' 查名, 申請, 延展, 補換發證書, 申請英文證明, 申請中文證明, 刊登廣告
            '93.9.30 modify by sonia
            'Case "101", "102", "103", "304", "309", "702":
            ' 查名, 申請, 延展, 補換發證書, 申請英文證明, 申請中文證明, 刊登廣告, 分割
            '2005/4/13 MODIFY BY SONIA 加入領土延伸
            'Case "101", "102", "103", "304", "309", "702", "308":
            ' 查名, 申請, 延展, 補換發證書, 領土延伸, 申請英文證明, 申請中文證明, 刊登廣告, 分割
            'Add By Sindy 2009/06/16 增加109.被異議續展
            Case "101", "102", "103", "104", "304", "309", "702", "308", "109":
               Set frmNext = frm020102_05
               bNext = True
            ' 變更, 更正, 減縮商品, 電話回覆  2007/6/7 加減縮商品 2009/12/3 加電話回覆
            Case "301", "302", "313", "209":
               Set frmNext = frm020102_07
               bNext = True
            ' 移轉
            Case "501":
               Set frmNext = frm020102_08
               bNext = True
            ' 授權, 再授權, 終止授權, 終止再授權 2009/10/14加徵求同意書724
            Case "502", "503", "504", "505", "724":
               Set frmNext = frm020102_09
               bNext = True
            'Modify By Cheng 2002/06/14
            ' 補正, 放棄專用權
'            Case "201":
            '2008/1/2 MODIFY BY SONIA 改203修正至此畫面
            Case "201", "206", "203":
               Set frmNext = frm020102_10
               bNext = True
            ' 延期
            Case "303":
               'Add By Cheng 2002/06/28
               frm020102_11.bln_From_Frm020102_01_BtnExt = False
               
               Set frmNext = frm020102_11
               bNext = True
            ' 自請撤回, 自請撤銷
            Case "306", "307":
               Set frmNext = frm020102_12
               bNext = True
            ' 設定質權, 撤銷設定質權
            Case "506", "507":
               Set frmNext = frm020102_13
               bNext = True
            ' 異議, 評定, 廢止, 評定專用權, 參加評定, 自評專用權, 禁止處分
            'modify by sonia 2014/11/4 +623部分廢止(龔2014/5/30郵件)
            'modify by sonia 2020/5/9 +627,629
            Case "601", "603", "605", "607", "608", "609", "616", "623", "627", "629":
               Set frmNext = frm020102_14
               bNext = True
            ' 申請意見書, 補充理由, 訴願, 再訴願, 行政訴訟, 參加行政訴訟, 再審之訴  92.7.11加 陳情
            '2011/11/4 modify by sonia 加210陳述意見書
            '2011/11/8 modify by sonia 加414再審之訴答辯
            'Modify By Sindy 2020/11/13 + 214.陳述聲明
            Case "202", "612", "401", "402", "403", "404", "405", "622", "210", "214", "414":
               Set frmNext = frm020102_15
               bNext = True
            ' 異議答辯, 評定答辯, 廢止答辯, 補充答辯, 參加被評定, 撤銷禁止處分, 修正, 參加訴訟, 第一期註冊費
'            Case "602", "604", "606", "613", "610", "617", "203", "407":
            '2008/1/2 MODIFY BY SONIA 改203修正至frm020102_10
            'modify by sonia 2014/11/4 +624部分廢止答辯(龔2014/5/30郵件)
            'Modified by Lydia 2023/10/13 +628 部分異議答辯, 630部分評定答辯
            Case "602", "604", "606", "613", "610", "617", "407", "715", "624", "628", "630":
               Set frmNext = frm020102_16
               bNext = True
            ' 補理由書
            Case "611":
               Set frmNext = frm020102_17
               bNext = True
            Case Else:
               Set frmNext = frm020102_16
               bNext = True
         End Select
   End Select
   
   ' 顯示下一個畫面
   'If IsObject(frmNext) = True Then
   If bNext = True Then
      frmNext.SetData 0, m_CP09, True
      '*********** 901121     nick
      If Me.Visible = True Then
         cmdQuery.Default = True
      End If
      '****************************
      Me.Hide
      frmNext.Show
      frmNext.QueryData
      
      ' 顯示商標基本資料的畫面
        'Modify By Cheng 2002/11/08
        '若案件性質為異議, 評定, 廢止時, 不要檢查是否要補基本檔資料
'      ShowMaintainForm m_CP09
        'edit by nick 2004/10/12  分割案也不用補
        'If StrCp10 <> "601" And StrCp10 <> "603" And StrCp10 <> "605" Then
        'Modify By Sindy 加 101申請,不開商標主檔,但要開申請人地址視窗
        'Modified by Lydia 2023/10/13 601+627, 603+629, 605+623
        'If strCP10 <> "101" And strCP10 <> "601" And strCP10 <> "603" And strCP10 <> "605" And strCP10 <> "308" Then
        If strCP10 <> "101" And strCP10 <> "601" And strCP10 <> "627" And strCP10 <> "603" And strCP10 <> "629" And strCP10 <> "605" And strCP10 <> "623" And strCP10 <> "308" Then
            ShowMaintainForm m_CP09
        'modify by sonia 2019/4/16 只要是A類新案件都要開申請人地址檢查
        'Else  frm020102_01
        End If
        'end 2019/4/16
            'Add By Sindy 2011/7/12 增加案件申請人地址視窗彈跳
            If (strCP01 = "T" Or strCP01 = "TF" Or strCP01 = "FCT") And Left(m_CP09, 1) = "A" And strCP31 = "Y" Then
               frm020102_23.Hide
               Set frm020102_23.UpForm = frmNext
               frm020102_23.m_CP09 = m_CP09
               'Me.Hide
               frm020102_23.QueryData
               frm020102_23.Show vbModal
            End If
            '2011/7/12 End
        'End If  cancel by sonia 2019/4/16
   End If
End Sub

' 顯示商標基本檔檔案畫面要求輸入
Private Sub DisplayFrm020501()
   frm020501.SetSystem 0
   frm020501.Show
End Sub

' 若為新案件且非新申請案且卷宗性質為"申請"時, 若商標基本檔的申請案號欄位是空白, 則先切換至商標基本資料維護
Private Function CheckJumpFrm020501() As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bShowFrm020501 As Boolean
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim strTM12 As String
   Dim strTM28 As String
   Dim strCP10 As String
   Dim strCP31 As String
   
   bShowFrm020501 = False
   
   strTM12 = Empty
   strCP10 = Empty
   strCP31 = Empty
   
   ' 查詢案件進度檔
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & m_CP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   
   If IsNull(rsTmp.Fields("CP01")) = False Then: strTM01 = rsTmp.Fields("CP01")
   If IsNull(rsTmp.Fields("CP02")) = False Then: strTM02 = rsTmp.Fields("CP02")
   If IsNull(rsTmp.Fields("CP03")) = False Then: strTM03 = rsTmp.Fields("CP03")
   If IsNull(rsTmp.Fields("CP04")) = False Then: strTM04 = rsTmp.Fields("CP04")
   ' 案件性質
   If IsNull(rsTmp.Fields("CP10")) = False Then
      If IsEmptyText(rsTmp.Fields("CP10")) = False Then
         strCP10 = rsTmp.Fields("CP10")
      End If
   End If
   ' 是否為新案件欄位
   If IsNull(rsTmp.Fields("CP31")) = False Then
      If IsEmptyText(rsTmp.Fields("CP31")) = False Then
         strCP31 = rsTmp.Fields("CP31")
      End If
   End If
   rsTmp.Close
   
   ' 查詢商標基本檔
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & strTM01 & "' AND " & _
                  "TM02 = '" & strTM02 & "' AND " & _
                  "TM03 = '" & strTM03 & "' AND " & _
                  "TM04 = '" & strTM04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount <= 0 Then: GoTo EXITSUB
   ' 卷宗性質
   If IsNull(rsTmp.Fields("TM28")) = False Then
      If IsEmptyText(rsTmp.Fields("TM28")) = False Then
         strTM28 = rsTmp.Fields("TM28")
      End If
   End If
   rsTmp.Close
   
   ' 判斷是否要顯示商標基本檔檔案維護的畫面
   If strTM28 = "1" Then
      If UCase(strCP31) = "Y" Then
         If strCP10 <> "101" Then
            bShowFrm020501 = True
         End If
      End If
   End If
      
   CheckJumpFrm020501 = False
EXITSUB:
   Set rsTmp = Nothing
End Function
' 檢查是否已選取資料
Private Function CheckDataValid() As Boolean
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   CheckDataValid = False
   If grdList.Rows <= 1 Then
      strTit = "檢核資料"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   If IsEmptyText(m_CP09) = True Then
      strTit = "檢核資料"
      strMsg = "請先選取資料"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      GoTo EXITSUB
   End If
   
   CheckDataValid = True
EXITSUB:
End Function

Public Sub SetData(ByVal nType As Integer, ByVal strData As String, Optional ByVal bClear As Boolean = False)
   ' 清除搜尋的Key
   If bClear = True Then
      m_TM01 = Empty
      m_TM02 = Empty
      m_TM03 = Empty
      m_TM04 = Empty
      m_CP09 = Empty
   End If
   
   Select Case nType
      ' 本所案號
      Case 0: m_TM01 = strData
      Case 1: m_TM02 = strData
      Case 2: m_TM03 = strData
      Case 3: m_TM04 = strData
   End Select
End Sub

' 更新查詢的方式由本所案號來查詢
Public Sub SetQueryFromTM()
   textTM01 = m_TM01
   textTM02 = m_TM02
   textTM03 = m_TM03
   textTM04 = m_TM04
   radio_Click 1
End Sub

Private Sub textCP09_GotFocus()
   InverseTextBox textCP09
End Sub

Private Sub textTM01_GotFocus()
   InverseTextBox textTM01
   CloseIme
End Sub
'2005/8/29 ADD BY SONIA
Private Sub textTM02_2_KeyPress(KeyAscii As Integer)
   cmdQuery.Default = True
End Sub

Private Sub textTM02_GotFocus()
   InverseTextBox textTM02
   CloseIme
End Sub

Private Sub textTM02_2_GotFocus()
   InverseTextBox textTM02_2
   CloseIme
End Sub
'2005/8/29 ADD BY SONIA
Private Sub textTM02_KeyPress(KeyAscii As Integer)
   cmdQuery.Default = True
End Sub

Private Sub textTM03_GotFocus()
   InverseTextBox textTM03
   CloseIme
End Sub

Private Sub textTM03_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textTM04_GotFocus()
   InverseTextBox textTM04
   CloseIme
End Sub

Private Sub textTM12_GotFocus()
   InverseTextBox textTM12
End Sub
