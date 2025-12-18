VERSION 5.00
Begin VB.Form frm04010401 
   BorderStyle     =   1  '單線固定
   Caption         =   "申請案號輸入"
   ClientHeight    =   1650
   ClientLeft      =   255
   ClientTop       =   990
   ClientWidth     =   5340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5340
   Begin VB.TextBox txtBillNo 
      Height          =   285
      Left            =   1224
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   1230
      Width           =   2385
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Index           =   1
      Left            =   4344
      TabIndex        =   5
      Top             =   48
      Width           =   800
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   400
      Index           =   0
      Left            =   3528
      TabIndex        =   4
      Top             =   48
      Width           =   800
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   1224
      TabIndex        =   3
      Top             =   855
      Width           =   1095
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   2808
      MaxLength       =   2
      TabIndex        =   2
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   2544
      MaxLength       =   1
      TabIndex        =   1
      Top             =   480
      Width           =   255
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1704
      MaxLength       =   6
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1224
      MaxLength       =   3
      TabIndex        =   6
      Text            =   "P"
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "帳單編號:"
      Height          =   180
      Left            =   120
      TabIndex        =   9
      Top             =   1282
      Width           =   765
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "來函收文日:"
      Height          =   180
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   948
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號:"
      Height          =   180
      Left            =   120
      TabIndex        =   7
      Top             =   525
      Width           =   768
   End
End
Attribute VB_Name = "frm04010401"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/16 改成Form2.0 (無)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit

Dim intWhere As Integer
'Add By Sindy 2016/9/21
Public m_strIR01 As String
Public m_strIR02 As String
Public m_strIR03 As String
Public m_strIR04 As String
Public m_strCP01 As String, m_strCP02 As String, m_strCP03 As String, m_strCP04 As String
Public m_RDate As String
Dim m_Done As Boolean
'2016/9/21 END

Private Sub cmdOK_Click(Index As Integer)
 Dim strTmp As String
 Dim bolChk As Boolean
   Select Case Index
      Case 0
         bolChk = False
         Text1_Validate bolChk
         If bolChk Then Exit Sub
         
         If Text2 = "" Then
            MsgBox "本所案號不可空白，請重新輸入 !", vbCritical
            Text2.SetFocus
            Exit Sub
         End If
         
         bolChk = False
         Text5_Validate bolChk
         If bolChk Then Exit Sub
         
         Text4_LostFocus
         strTmp = Text1 & Text2 & Text3 & Text4
         Select Case Text1.Text
            Case "P"
               strExc(0) = "SELECT PA01,PA02,PA03,PA04 FROM PATENT WHERE " & ChgPatent(strTmp)
            Case "PS"
               strExc(0) = "SELECT SP01,SP02,SP03,SP04 FROM SERVICEPRACTICE WHERE " & ChgService(strTmp)
         End Select
         intI = 0
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '92.10.22 MODIFY BY SONIA
            'strExc(0) = "select CP05,CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN ('101','102','103','104','105','301','302','303','304','305','306','307') AND " & _
            '   "CP27 IS NOT NULL and cp05= (" & _
            '   "select MAX(CP05) FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN ('101','102','103','104','105','301','302','303','304','305','306','307') AND " & _
            '   "CP27 IS NOT NULL) "
            '93.4.28 modify by sonia 加 積體電路佈局 117
            'strExc(0) = "select CP05,CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN ('101','102','103','104','105','109','110','112','301','302','303','304','305','306','307') AND " & _
            '   "CP27 IS NOT NULL and cp05= (" & _
            '   "select MAX(CP05) FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN ('101','102','103','104','105','109','110','112','301','302','303','304','305','306','307') AND " & _
            '   "CP27 IS NOT NULL) "
            
            'Modify by Morgan 2005/10/18 加提申日CP47,PA09
'            strExc(0) = "select CP05,CP09 FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
'               " AND CP10 IN ('101','102','103','104','105','109','110','112','117','301','302','303','304','305','306','307') AND " & _
'               "CP27 IS NOT NULL and cp05= (" & _
'               "select MAX(CP05) FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
'               " AND CP10 IN ('101','102','103','104','105','109','110','112','117','301','302','303','304','305','306','307') AND " & _
'               "CP27 IS NOT NULL) "
'            '92.10.22 END
            '2006/5/9 MODIFY BY SONIA 新申請案未收達不可輸入申請案號
            'strExc(0) = "select CP05,CP09,CP47,PA09 FROM CASEPROGRESS,PATENT WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN ('101','102','103','104','105','109','110','112','117','301','302','303','304','305','306','307') " & _
            '   " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND CP27 IS NOT NULL and cp05= (" & _
            '   "select MAX(CP05) FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN ('101','102','103','104','105','109','110','112','117','301','302','303','304','305','306','307') AND " & _
            '   "CP27 IS NOT NULL) "
            'Modify by Morgan 2007/1/17 加判斷CP31='Y'因為大陸案會有多收翻譯201的，且若沒有CP31='Y'的表示資料有錯應更正
            'strExc(0) = "select CP05,CP09,CP47,PA09,CP46 FROM CASEPROGRESS,PATENT WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN (" & CaseMapIn & ",117,301,302,303,304,305,306) " & _
            '   " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND CP27 IS NOT NULL and cp05= (" & _
            '   "select MAX(CP05) FROM CASEPROGRESS WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN (" & CaseMapIn & ",117,301,302,303,304,305,306) AND " & _
            '   "CP27 IS NOT NULL)"
            'Modify by Morgan 2007/9/5 改請的CP31<>'Y'所以改排除翻譯就好
            'strExc(0) = "select CP05,CP09,CP47,PA09,CP46 FROM CASEPROGRESS,PATENT WHERE " & ChgCaseprogress(strTmp) & _
            '   " AND CP10 IN (" & CaseMapIn & ",117,301,302,303,304,305,306) AND CP31='Y'" & _
            '   " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND CP27 IS NOT NULL ORDER BY 1 DESC"
            '2010/4/6 modify by sonia 加香港111(P-086224)
            'Modified by Morgan 2013/1/14 +308,309
            'Modified by Morgan 2014/2/14 改以發文日排序(收文日可能會同一天 ex.P-107174)
            strExc(0) = "select CP05,CP09,CP47,PA09,CP46 FROM CASEPROGRESS,PATENT WHERE " & ChgCaseprogress(strTmp) & _
               " AND CP10 IN (" & CaseMapIn & ",117,301,302,303,304,305,306,308,309,111) AND CP10<>'201'" & _
               " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND CP27 IS NOT NULL ORDER BY cp27 DESC"
            'END 2007/1/17
            '2006/5/9 END
            '2005/10/18 END
            
       'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
        If FMP2open = True Then
           If PUB_FMPtoCheck(0, 1, Pub_strUserST05, Text1, Text2, Text3, Text4) = False Then
             'Text1 = "P": Text2 = "": Text3 = "": Text4 = "" '無權限清空
             Text2.SetFocus
             Exit Sub
           ElseIf FMP2openSQL <> "" Then
             strExc(0) = "select CP05,CP09,CP47,PA09,CP46 FROM CASEPROGRESS f0,PATENT WHERE " & ChgCaseprogress(strTmp) & _
               " AND CP10 IN (" & CaseMapIn & ",117,301,302,303,304,305,306,308,309,111) AND CP10<>'201'" & FMP2openSQL & _
               " AND PA01(+)=CP01 AND PA02(+)=CP02 AND PA03(+)=CP03 AND PA04(+)=CP04 AND CP27 IS NOT NULL ORDER BY cp27 DESC"
           End If
        'Added by Lydia 2019/09/10 寰華案控制輸入帳單、已提申、發證書輸入，P的程序不能輸入並出現訊息告知USER「此案為FCP自行連繫，請交FCP程序處理」。
        ElseIf Pub_StrUserSt03 <> "M51" Then
           If PUB_FMPtoCheck(1, 2, Pub_strUserST05, Text1, Text2, Text3, Text4) = True Then
                MsgBox "此案為FCP自行連繫，請交FCP程序處理！", vbCritical, "寰華案控制輸入"
                Exit Sub
           End If
        'end 2019/09/10
        End If
        
            intI = 0
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Add by Morgan 2005/10/18 重複輸入申請日時提醒
               If RsTemp.Fields("PA09") <> "000" Then
                  If Not IsNull(RsTemp.Fields("CP47")) Then
                     If MsgBox("本案已輸過提申日，是否要再次輸入？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbNo Then
                        Exit Sub
                     End If
                  End If
               End If
               
               '2006/5/9 ADD BY SONIA
               If RsTemp.Fields("PA09") <> "000" Then
                  If IsNull(RsTemp.Fields("CP46")) Then
                     MsgBox "本案尚未輸入收達日，不可輸入申請案號 ! ", vbCritical
                     Exit Sub
                  End If
               End If
               
               'Add By Sindy 2017/12/27
               If m_strIR01 <> "" Then
                  If m_strCP01 & m_strCP02 & m_strCP03 & m_strCP04 <> Text1 & Text2 & Text3 & Text4 Then
                     MsgBox "信件輸入必須與信件本所案號(" & m_strCP01 & "-" & m_strCP02 & "-" & m_strCP03 & "-" & m_strCP04 & ")一致！"
                     Exit Sub
                  End If
               End If
               '2017/12/27 END
               
               'Added by Morgan 2021/12/16
               '配合改Form2.0Unload可能沒有下此指令改在呼叫前執行以避免前次變數值殘留
               If PUB_CheckFormExist("frm040101_1") = False Then
                  Set frm040101_1 = Nothing
               End If
               'end 2021/12/16
   
               frm04010402.SetData strTmp, RsTemp.Fields("CP09")
               frm04010402.Show
               frm04010402.SetDefault
               'Add By Sindy 2016/9/21
               frm04010402.m_strIR01 = m_strIR01
               frm04010402.m_strIR02 = m_strIR02
               frm04010402.m_strIR03 = m_strIR03
               frm04010402.m_strIR04 = m_strIR04
               '2016/9/21 END
               Me.Hide
            Else
               TextInverse Text2
            End If
         Else
            TextInverse Text2
         End If
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Activate()
   'Added by Sindy 2016/9/21
   If m_strIR01 <> "" And m_Done = False Then
      Text1.Text = m_strCP01
      Text2.Text = m_strCP02
      Text3.Text = m_strCP03
      Text4.Text = m_strCP04
      Text5.Text = m_RDate
      cmdOK(0).Value = True
      m_Done = True
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）" 'Add By Sindy 2017/12/27
   End If
   '2016/9/21 END
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   intWhere = 國內
   Text5.Text = strSrvDate(2)
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
    FMP2open = PUB_FMPtoCheck(1, 0, Pub_strUserST05)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm04010401 = Nothing
End Sub

Private Sub Text1_GotFocus()
   InverseTextBox Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Text1.Text <> "P" And Text1.Text <> "PS" Then
      MsgBox "本所案號不正確，請重新輸入 !", vbCritical
      Cancel = True
      InverseTextBox Text1
   End If
End Sub

Private Sub Text2_GotFocus()
   InverseTextBox Text2
End Sub

Private Sub Text3_GotFocus()
   InverseTextBox Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text4_GotFocus()
   InverseTextBox Text4
End Sub

Private Sub Text4_LostFocus()
   If Text3 = "" Then Text3 = "0"
   If Text4 = "" Then Text4 = "00"
End Sub

Private Sub Text5_GotFocus()
   TextInverse Text5
End Sub

Private Sub Text5_Validate(Cancel As Boolean)
   If Text5 = "" Then
      MsgBox "來函收文日不可空白，請重新輸入 !", vbCritical
      Cancel = True
   Else
      If ChkDate(Text5) Then
         If Val(Text5) > Val(strSrvDate(2)) Then
            MsgBox "來函收文日不可大於系統日 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse Text5
End Sub
