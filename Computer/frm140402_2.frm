VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140402_2 
   BorderStyle     =   1  '單線固定
   Caption         =   "潛在客戶轉客戶或代理人作業"
   ClientHeight    =   5736
   ClientLeft      =   660
   ClientTop       =   636
   ClientWidth     =   8100
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5736
   ScaleWidth      =   8100
   Begin VB.Frame Frame1 
      Caption         =   "設定 代理人 來所原因"
      Height          =   2000
      Left            =   3600
      TabIndex        =   30
      Top             =   3570
      Width           =   4400
      Begin VB.TextBox txtXYS02 
         Height          =   315
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   33
         Top             =   600
         Width           =   1000
      End
      Begin VB.ComboBox cboSource 
         Height          =   300
         ItemData        =   "frm140402_2.frx":0000
         Left            =   960
         List            =   "frm140402_2.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   31
         Top             =   225
         Width           =   3350
      End
      Begin MSForms.TextBox txtXYS03 
         Height          =   900
         Left            =   570
         TabIndex        =   37
         Top             =   990
         Width           =   3700
         VariousPropertyBits=   -1466941413
         MaxLength       =   1000
         ScrollBars      =   2
         Size            =   "6526;1587"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   " 其他     說明："
         Height          =   492
         Index           =   10
         Left            =   30
         TabIndex        =   36
         Top             =   996
         Width           =   588
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "介紹者編號："
         Height          =   180
         Index           =   9
         Left            =   60
         TabIndex        =   35
         Top             =   600
         Width           =   1080
      End
      Begin MSForms.Label LblSourceN 
         Height          =   285
         Left            =   2280
         TabIndex        =   34
         Top             =   600
         Width           =   2000
         VariousPropertyBits=   27
         Caption         =   "LblSourceN"
         Size            =   "3528;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "來所原因："
         Height          =   180
         Index           =   7
         Left            =   60
         TabIndex        =   32
         Top             =   225
         Width           =   900
      End
   End
   Begin VB.TextBox txtCU153 
      Height          =   285
      Left            =   420
      MaxLength       =   1
      TabIndex        =   29
      Top             =   120
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.OptionButton Option1 
      Caption         =   "新代理人"
      Height          =   180
      Index           =   3
      Left            =   3630
      TabIndex        =   9
      Top             =   2880
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "舊代理人關係企業："
      Height          =   180
      Index           =   4
      Left            =   3630
      TabIndex        =   8
      Top             =   3240
      Width           =   1950
   End
   Begin VB.OptionButton Option1 
      Caption         =   "舊客戶關係企業："
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   7
      Top             =   3240
      Width           =   1750
   End
   Begin VB.OptionButton Option1 
      Caption         =   "新客戶"
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   6
      Top             =   2880
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消(&U)"
      Height          =   400
      Index           =   1
      Left            =   7110
      TabIndex        =   2
      Top             =   70
      Width           =   800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "確定(&O)"
      Height          =   400
      Index           =   0
      Left            =   6270
      TabIndex        =   1
      Top             =   70
      Width           =   800
   End
   Begin VB.TextBox TextOldNo 
      Height          =   264
      Left            =   2040
      MaxLength       =   8
      TabIndex        =   0
      Top             =   3600
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   13
      Left            =   960
      TabIndex        =   28
      Top             =   5280
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   12
      Left            =   960
      TabIndex        =   27
      Top             =   4950
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   11
      Left            =   960
      TabIndex        =   26
      Top             =   4725
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   10
      Left            =   960
      TabIndex        =   25
      Top             =   4485
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   9
      Left            =   960
      TabIndex        =   24
      Top             =   4260
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   8
      Left            =   960
      TabIndex        =   23
      Top             =   3960
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   7
      Left            =   960
      TabIndex        =   22
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   6
      Left            =   960
      TabIndex        =   21
      Top             =   2070
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   5
      Left            =   960
      TabIndex        =   20
      Top             =   1840
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   4
      Left            =   960
      TabIndex        =   19
      Top             =   1610
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   3
      Left            =   960
      TabIndex        =   18
      Top             =   1380
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   2
      Left            =   960
      TabIndex        =   17
      Top             =   1080
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Label2"
      Height          =   180
      Index           =   1
      Left            =   1680
      TabIndex        =   16
      Top             =   720
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "PS：舊名稱資料會一併轉出"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   0
      Left            =   3600
      TabIndex        =   15
      Top             =   720
      Width           =   2160
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "舊客戶/代理人編號："
      Height          =   180
      Index           =   8
      Left            =   360
      TabIndex        =   14
      Top             =   3600
      Width           =   1665
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "潛在客戶編號："
      Height          =   180
      Index           =   17
      Left            =   360
      TabIndex        =   13
      Top             =   720
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日文："
      Height          =   180
      Index           =   6
      Left            =   360
      TabIndex        =   12
      Top             =   2400
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文："
      Height          =   180
      Index           =   5
      Left            =   360
      TabIndex        =   11
      Top             =   1380
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "中文："
      Height          =   180
      Index           =   3
      Left            =   360
      TabIndex        =   10
      Top             =   1080
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "中文："
      Height          =   180
      Index           =   1
      Left            =   360
      TabIndex        =   5
      Top             =   3960
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "英文："
      Height          =   180
      Index           =   2
      Left            =   360
      TabIndex        =   4
      Top             =   4260
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日文："
      Height          =   180
      Index           =   4
      Left            =   360
      TabIndex        =   3
      Top             =   5280
      Width           =   540
   End
End
Attribute VB_Name = "frm140402_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2022/1/7 改成Form2.0 (無)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/26 員工編號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
'2008/12/4 add by sonia
Option Explicit
Public m_PCU37 As String, m_PCU38 As String 'Added by Lydia 2018/05/16 潛在客戶-開發日期,開發人員
Public m_PCU47 As String, m_PCU49 As String 'Added by Lydia 2020/08/27  國外部關聯企業：潛在客戶的原設定、關係
Dim bCancel As Boolean 'Add by Amy 2023/05/08
Dim stMsg As String 'Add by Amy 2024/11/29

'Add by Amy 2023/05/08
Private Sub cboSource_Click()
    If cboSource = MsgText(601) Then Exit Sub
   
    'Modify by Amy 2024/11/29 改成共用函數,避免未改到
    txtXYS02.Text = "": LblSourceN.Caption = ""
    Call Pub_SetCboComeSource(9, Me.Name, cboSource, , txtXYS02, txtXYS03)
    
End Sub

Private Sub Form_Load()
 
   MoveFormToCenter Me
   
   TextOldNo.Text = ""
   TextOldNo_Validate True
   'Add by Amy 2023/05/08 來所原因(原:代理人來源)
   'Modify by Amy 2024/11/29 改為共用函數,避免未改到
   cboSource.ListIndex = -1
   LblSourceN.Caption = ""
   Call Pub_SetCboComeSource(0, Me.Name, cboSource)
   'end 2023/05/08
End Sub

Private Sub Command1_Click(Index As Integer)
Dim m_TranNo As String, nFrm As Form 'Add by Amy 2024/01/22 轉後編號/表單

   Select Case Index
      Case 0 '確定
         Screen.MousePointer = vbHourglass
         
         If TxtValidate() = True Then
            If FormSave(m_TranNo) = False Then Screen.MousePointer = vbDefault: Exit Sub
            PUB_SendMailCache 'Added by Lydia 2018/05/16
            MsgBox "資料已轉成功,稍後將開啟 " & IIf(Left(m_TranNo, 1) = "Y", "代理人", "客戶") & " 維護畫面" 'Add by Amy 2024/01/22
         Else
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         
         Screen.MousePointer = vbDefault
         '結束回前畫面帶出下一筆資料
         Unload frm140402_2
         Set frm140402_2 = Nothing
         'Modify by Amy 2024/01/22 切換至客戶或代理人資料維護-陳金蓮
         If Left(m_TranNo, 1) = 代理人編號 Then
            Call frm050705.SetParent(frm140402, m_TranNo)
            frm050705.Show
         Else
            Call frm140401.SetParent(frm140402, m_TranNo)
            frm140401.Show
         End If
         'frm140402.AfterTransfer
         'end 2024/01/22
         
      Case 1 '取消
         If MsgBox("你並未存檔，確定離開 ?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
         Unload frm140402_2
         Set frm140402_2 = Nothing
         
   End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm140402_2 = Nothing
End Sub

'Add by Amy 2023/05/08
Private Sub Option1_Click(Index As Integer)
    If Frame1.Visible = False Then Exit Sub
    
    Frame1.Enabled = False
    '選 新代理人 Or 舊代理人關係企業
    If Index = 3 Or Index = 4 Then
        Frame1.Enabled = True
    End If
End Sub

Private Sub TextOldNo_GotFocus()
   CloseIme
   TextInverse TextOldNo
End Sub

Private Sub TextOldNo_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TextOldNo_Validate(Cancel As Boolean)

   Label2(8) = ""
   Label2(9) = ""
   Label2(10) = ""
   Label2(11) = ""
   Label2(12) = ""
   Label2(13) = ""
   
   If TextOldNo <> "" Then
      If Option1(2).Visible = True And Mid(TextOldNo, 1, 1) <> "X" Then
         ShowMsg "舊客戶編號, 請輸入 X 字頭編號 !"
         TextOldNo.SetFocus
         Cancel = True
      ElseIf Option1(4).Visible = True And Mid(TextOldNo, 1, 1) <> "Y" Then
         ShowMsg "舊代理人編號, 請輸入 Y 字頭編號 !"
         TextOldNo.SetFocus
         Cancel = True
      End If
      
      If Cancel = False Then
         TextOldNo = TextOldNo + String(TextOldNo.MaxLength - Len(TextOldNo), "0")
         strExc(0) = "SELECT CU04,CU05,CU88,CU89,CU90,CU06 FROM CUSTOMER WHERE CU01='" & TextOldNo & "' AND CU02='0' UNION " & _
                     "SELECT FA04,FA05,FA63,FA64,FA65,FA06 FROM FAGENT   WHERE FA01='" & TextOldNo & "' AND FA02='0' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Label2(8) = "" & RsTemp.Fields(0)
            Label2(9) = "" & RsTemp.Fields(1)
            Label2(10) = "" & RsTemp.Fields(2)
            Label2(11) = "" & RsTemp.Fields(3)
            Label2(12) = "" & RsTemp.Fields(4)
            Label2(13) = "" & RsTemp.Fields(5)
         Else
            ShowMsg "無此舊客戶或代理人編號 !"
            TextOldNo.SetFocus
            Cancel = True
         End If
      End If
   End If
   
   If Cancel = True Then TextInverse TextOldNo
End Sub

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean, ii As Integer, jj As Integer
Dim stMsg As String 'Add by Amy 2024/11/29

   TxtValidate = False

   If (Option1(2).Value = True Or Option1(4).Value = True) Then
      If TextOldNo = "" Then
         ShowMsg "選擇舊客戶或代理人關係企業時, 請輸入舊客戶/代理人編號 !"
         TextOldNo.SetFocus
         Exit Function
      Else
         TextOldNo_Validate Cancel
         If Cancel = True Then Exit Function
      End If
   '新客戶 or 新代理人
   ElseIf (Option1(1).Value = True Or Option1(3).Value = True) Then
      TextOldNo = ""
      Label2(8) = ""
      Label2(9) = ""
      Label2(10) = ""
      Label2(11) = ""
      Label2(12) = ""
      Label2(13) = ""
   End If
   'Add by Amy 2023/05/08 新代理人 or 舊代理人關係企業,來所原因 必填
   Cancel = False
   If (Option1(3) = True Or Option1(4) = True) And Frame1.Visible = True Then
        'Modify by Amy 2024/11/29 改成共用函數,避免有未改到
        stMsg = ChkXYSourceReason(0, Me.Name, 1, cboSource, txtXYS02, , , , txtXYS03)
      If stMsg <> MsgText(601) Then
         MsgBox stMsg, vbInformation
         If InStr(stMsg, "來所原因 不可為空") > 0 Then
            cboSource.SetFocus
         ElseIf InStr(stMsg, "介紹者編號") > 0 Then
            txtXYS02.SetFocus
         ElseIf InStr(stMsg, "其他說明") > 0 Then
            txtXYS03.SetFocus
         End If
         Exit Function
      End If
      'end 2024/11/11
   End If
   
   TxtValidate = True
   
End Function

'Modify by Amy 2024/01/022 +m_NewNo
Private Function FormSave(Optional ByRef m_NewNo As String) As Boolean
Dim strTxt(1 To 10) As String
'Dim m_NewNo As String 'Mark by Amy 2024/01/22
Dim Newno As Variant
Dim stXYSMemo As String 'Add by Amy 2024/11/29
Dim stFixPCU As String 'Add by Amy 2025/11/18

   FormSave = False
   m_NewNo = ""
   '先依選擇找出新編號
   If Option1(1).Value = True Then             '新客戶
      If ClsPDGetAutoNumber("X", m_NewNo, True, False) Then
         m_NewNo = "X" + Right(m_NewNo, 5)
         m_NewNo = m_NewNo & String(8 - Len(m_NewNo), "0")
      Else
         ShowMsg "讀取自動編號檔錯誤，請洽系統管理者 !"
         Exit Function
      End If
   ElseIf Option1(2).Value = True Then         '舊客戶關係企業
      strExc(0) = "SELECT SUBSTR(MAX(CU01),2,7) FROM CUSTOMER WHERE SUBSTR(CU01,1,6)='" & Mid(TextOldNo, 1, 6) & "' AND CU02='0'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Newno = RsTemp.Fields(0) + 1
         m_NewNo = "X" & Newno
         m_NewNo = m_NewNo & String(8 - Len(m_NewNo), "0")
      Else
         ShowMsg "讀取客戶檔錯誤，請洽系統管理者 !"
         Exit Function
      End If
   ElseIf Option1(3).Value = True Then         '新代理人
      If ClsPDGetAutoNumber("Y", m_NewNo, True, False) Then
         m_NewNo = "Y" + Right(m_NewNo, 5)
         m_NewNo = m_NewNo & String(8 - Len(m_NewNo), "0")
      Else
         ShowMsg "讀取自動編號檔錯誤，請洽系統管理者 !"
         Exit Function
      End If
   ElseIf Option1(4).Value = True Then         '舊代理人關係企業
      strExc(0) = "SELECT SUBSTR(MAX(FA01),2,7) FROM FAGENT   WHERE SUBSTR(FA01,1,6)='" & Mid(TextOldNo, 1, 6) & "' AND FA02='0' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modified by Lydia 2020/05/27 判斷第7和第8碼是否為數字; ex.  R00595轉Y30117, 因為關係企業編號最大為Y30117B30
         'Newno = RsTemp.Fields(0) + 1
         If Right("" & RsTemp.Fields(0), 2) = Format(Val(Right("" & RsTemp.Fields(0), 2)), "00") Then
             Newno = Mid(RsTemp.Fields(0), 1, 5) & Format(Val(Mid(RsTemp.Fields(0), 6, 2)) + 1, "00")
         Else
            strExc(1) = Mid("" & RsTemp.Fields(0), 6, 1) '第7碼
            strExc(2) = Mid("" & RsTemp.Fields(0), 7, 1) '第8碼
            If InStr("0123456789", strExc(2)) > 0 Then
                Newno = Mid(RsTemp.Fields(0), 1, 5) & strExc(1) & IIf(strExc(2) = "9", "A", Val(strExc(2)) + 1)
            Else  '第8碼超過9，從A開頭
                Newno = Mid(RsTemp.Fields(0), 1, 5) & strExc(1) & Chr(Asc(strExc(2)) + 1)
            End If
         End If
         'end 2020/05/27
         m_NewNo = "Y" & Newno
         m_NewNo = m_NewNo & String(8 - Len(m_NewNo), "0")
      Else
         ShowMsg "讀取代理人檔錯誤，請洽系統管理者 !"
         Exit Function
      End If
   End If
    
On Error GoTo ErrorHandler
   cnnConnection.BeginTrans
    
   stFixPCU = ",'^原潛在客戶編號:'||PCU01||PCU02||DECODE(PCU19,'',NULL,';網址:')||PCU19"
   '將潛在客戶及其舊名稱都轉出
   'modify by sonia 2013/6/26 加入是否寄發專利雙週報PCU48
   If Option1(1).Value = True Or Option1(2).Value = True Then       '轉入客戶
      'Modify By Sindy 2014/7/10 +CU153
      'Modify by Amy 2024/11/29 國外潛在客戶 加 來所原因(非必填,傳入此畫面時畫面為隱藏),其資料轉入時寫於備註
      If cboSource <> MsgText(601) Then
         stXYSMemo = "||';來所原因:" & cboSource & "'"
      End If
      'Modify by Amy 2025/11/18 原R編號之備註改位置
      strTxt(1) = stFixPCU & stXYSMemo & "||Decode(PCU40,null,'',';備註:'||PCU40)||'^' "
'      strTxt(1) = "INSERT INTO CUSTOMER (CU01,CU02,CU05,CU88,CU89,CU90,CU06,CU04,CU10,CU16,CU17,CU18,CU19,CU22,CU20,CU79,CU24,CU25,CU26,CU27,CU28,CU102," & _
'                  "CU29,CU23,CU87,CU65,CU66,CU67,CU68,CU69,CU32,CU132,CU64,CU14,CU129,CU80,CU84,CU85,CU86,CU145,CU153) " & _
'                  "(SELECT '" & m_NewNo & "',PCU02,PCU03,PCU04,PCU05,PCU06,PCU07,PCU08,PCU09,PCU13,PCU14,PCU15,PCU16,PCU17,PCU18,PCU40||';原潛在客戶編號:'||PCU01||PCU02||DECODE(PCU19,'',NULL,';網址:')||PCU19" & stXYSMemo & ",PCU20,PCU21,PCU22,PCU23,PCU24,PCU25," & _
'                  "PCU26,PCU27,PCU28,PCU29,PCU30,PCU31,PCU32,PCU33,PCU34,PCU35,PCU36,PCU37,PCU38,PCU39,PCU44,PCU45,PCU46,PCU48," & CNULL(txtCU153) & " FROM PotCustomer WHERE PCU01='" & Left(Label2(1), 8) & "')"
      strTxt(1) = "INSERT INTO CUSTOMER (CU01,CU02,CU05,CU88,CU89,CU90,CU06,CU04,CU10,CU16,CU17,CU18,CU19,CU22,CU20,CU79,CU24,CU25,CU26,CU27,CU28,CU102," & _
                  "CU29,CU23,CU87,CU65,CU66,CU67,CU68,CU69,CU32,CU132,CU64,CU14,CU129,CU80,CU84,CU85,CU86,CU145,CU153) " & _
                  "(SELECT '" & m_NewNo & "',PCU02,PCU03,PCU04,PCU05,PCU06,PCU07,PCU08,PCU09,PCU13,PCU14,PCU15,PCU16,PCU17,PCU18" & strTxt(1) & ",PCU20,PCU21,PCU22,PCU23,PCU24,PCU25," & _
                  "PCU26,PCU27,PCU28,PCU29,PCU30,PCU31,PCU32,PCU33,PCU34,PCU35,PCU36,PCU37,PCU38,PCU39,PCU44,PCU45,PCU46,PCU48," & CNULL(txtCU153) & " FROM PotCustomer WHERE PCU01='" & Left(Label2(1), 8) & "')"
      'end 2024/11/29
   ElseIf Option1(3).Value = True Or Option1(4).Value = True Then   '轉入代理人且性質FA76='A'
      'Modify by Amy 2023/05/08 +fa127 來所原因
      strTxt(10) = ""
      strExc(9) = stFixPCU & "||DECODE(PCU17,'',NULL,';行動電話:')||PCU17||';備註:'||PCU40||'^' "
      strTxt(1) = "FA01,FA02,FA05,FA63,FA64,FA65,FA06,FA04,FA10,FA12,FA13,FA14,FA15,FA16,FA29,FA18,FA19,FA20,FA21,FA22,FA70," & _
                  "FA23,FA17,FA55,FA32,FA33,FA34,FA35,FA36,FA24,FA97,FA31,FA11,FA94,FA69,FA49,FA50,FA51,FA76,FA100,FA123,FA127"
      strTxt(10) = ",'" & Left(cboSource, 2) & "'"
      'Modified by Lydia 2018/06/28 +PCU50,FA123 是否同意歐盟通用資料保護規範(GDPR)
'      strTxt(1) = "INSERT INTO FAGENT (" & strTxt(1) & ") " & _
'                  "(SELECT '" & m_NewNo & "',PCU02,PCU03,PCU04,PCU05,PCU06,PCU07,PCU08,PCU09,PCU13,PCU14,PCU15,PCU16,PCU18,PCU40||';原潛在客戶編號:'||PCU01||PCU02||DECODE(PCU19,'',NULL,';網址:')||PCU19||DECODE(PCU17,'',NULL,';行動電話:')||PCU17,PCU20,PCU21,PCU22,PCU23,PCU24,PCU25," & _
'                  "PCU26,PCU27,PCU28,PCU29,PCU30,PCU31,PCU32,PCU33,PCU34,PCU35,PCU36,PCU37,PCU38,PCU39,PCU44,PCU45,PCU46,'A',PCU48,PCU50" & strTxt(10) & " FROM PotCustomer WHERE PCU01='" & Left(Label2(1), 8) & "')"
      strTxt(1) = "INSERT INTO FAGENT (" & strTxt(1) & ") " & _
                  "(SELECT '" & m_NewNo & "',PCU02,PCU03,PCU04,PCU05,PCU06,PCU07,PCU08,PCU09,PCU13,PCU14,PCU15,PCU16,PCU18" & strExc(9) & ",PCU20,PCU21,PCU22,PCU23,PCU24,PCU25," & _
                  "PCU26,PCU27,PCU28,PCU29,PCU30,PCU31,PCU32,PCU33,PCU34,PCU35,PCU36,PCU37,PCU38,PCU39,PCU44,PCU45,PCU46,'A',PCU48,PCU50" & strTxt(10) & " FROM PotCustomer WHERE PCU01='" & Left(Label2(1), 8) & "')"
      'end 2023/05/08
   End If
   Pub_SeekTbLog strTxt(1)
   cnnConnection.Execute strTxt(1)
   
   'Modify by Amy 2024/11/29 先更XYS02資料,再確認XYNoSource 是否已有資料,再依狀況修改其資料
   'Add by Amy 2023/05/08 +代理人來源
'   If txtXYS02 <> MsgText(601) Or txtXYS03 <> MsgText(601) Then
'        Call SaveXYNoSource(1, Me.Name, m_NewNo, txtXYS02, txtXYS03)
'    End If
   stMsg = SaveXYNoSource(4, Me.Name, m_NewNo, , , , , Left(Label2(1), 8))
   If Len(stMsg) > 1 Then
      GoTo ErrorHandler
   End If
   stMsg = ""
   '確認XYNoSource 資料再視狀況改資料,因 國外潛在客戶 加 來所原因,於此頁面可調整資料
   If ExistCheck("XYNoSource", "XYS01", Left(Label2(1), 8), "", False) = False Then
      stMsg = SaveXYNoSource(1, Me.Name, m_NewNo, txtXYS02, txtXYS03, Left(cboSource, 2))
   Else
      stMsg = SaveXYNoSource(2, Me.Name, m_NewNo, txtXYS02, txtXYS03, Left(cboSource, 2), cboSource.Tag, Left(Label2(1), 8))
   End If
   If Len(stMsg) > 1 Then
      GoTo ErrorHandler
   End If
   stMsg = ""
   'end 2024/11/29
   
   '原潛在客戶檔刪除
   strTxt(2) = "DELETE PotCustomer WHERE PCU01='" & Left(Label2(1), 8) & "'"
   Pub_SeekTbLog strTxt(2)
   cnnConnection.Execute strTxt(2)
   'Added by Lydia 2024/05/14 聯絡人相片
   strExc(0) = "select ibf01,ibf02,ibf03,ibf04,ibf05,pcc01,pcc02 from potcustcont,imgbytefile where pcc01='" & Left(Label2(1), 8) & "' and pcc01||pcc02=ibf01||ibf02||ibf03 and ibf04='00' and ibf05='3' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         strSql = "Update ImgByteFile Set IBF01='" & Pub_GetPCCtoIBF(m_NewNo, RsTemp.Fields("pcc02"), "1") & "',IBF02='" & Pub_GetPCCtoIBF(m_NewNo, RsTemp.Fields("pcc02"), "2") & "' " & _
                  ",IBF03='" & Pub_GetPCCtoIBF(m_NewNo, RsTemp.Fields("pcc02"), "3") & "' Where ibf01='" & RsTemp.Fields("ibf01") & "' and ibf02='" & RsTemp.Fields("ibf02") & "' " & _
                  "and ibf03='" & RsTemp.Fields("ibf03") & "' and ibf04='00' and ibf05 = '3' "
         cnnConnection.Execute strSql
         RsTemp.MoveNext
      Loop
   End If
   'end 2024/05/14
   '其所有聯絡人資料也轉出
   strTxt(3) = "UPDATE PotCustCont SET PCC01='" & m_NewNo & "' WHERE PCC01='" & Left(Label2(1), 8) & "'"
   Pub_SeekTbLog strTxt(3)
   cnnConnection.Execute strTxt(3)
   '其所有往來記錄資料也轉出
   strTxt(4) = "UPDATE ContactRecord SET CR03='" & m_NewNo & "'||SUBSTR(CR03,9,1) WHERE SUBSTR(CR03,1,8)='" & Left(Label2(1), 8) & "'"
   Pub_SeekTbLog strTxt(4)
   cnnConnection.Execute strTxt(4)
   'Added by Lydia 2022/03/28 DHL輸入資料
    strSql = "UPDATE  DHL_INPUT_DATA SET DID01 = '" & m_NewNo & "', DID02 = '" & Mid(Label2(1), 9, 1) & "' " & _
             "WHERE DID01 = '" & Left(Label2(1), 8) & "' AND DID02 = '" & Mid(Label2(1), 9, 1) & "' "
    cnnConnection.Execute strSql
    
   'Added by Lydia 2020/08/27 預設國外關聯企業
   'Remark by Lydia 2021/01/06 潛在客戶不使用關聯企業設定
'   If strSrvDate(1) >= 國外部關聯企業啟用日 And m_PCU47 <> "" And m_PCU49 <> "" Then
'       strSql = "INSERT INTO FRELATION (FR01,FR02,FR03,FR04,FR05,FR06,FR07) " & _
'                "VALUES ('" & m_NewNo & "','" & m_PCU47 & "','" & m_PCU49 & "','原潛在客戶編號:" & Label2(1).Caption & "','" & strUserNum & "'," & strSrvDate(1) & "," & Left(Format(ServerTime, "000000"), 4) & ")"
'       Pub_SeekTbLog strSql
'       cnnConnection.Execute strSql, intI
'   End If
'   'end 2020/08/27
   'end 2021/01/06
   
   'Added by Lydia 2018/05/16 當X、Y編號是從R編號轉成時，自動以email一併通知開發人員
    strExc(1) = "英：" & Label2(3).Caption & IIf(Label2(4).Caption <> "", vbCrLf & "　　" & Label2(4).Caption, "") & _
                                IIf(Label2(5).Caption <> "", vbCrLf & "　　" & Label2(5).Caption, "") & _
                                IIf(Label2(6).Caption <> "", vbCrLf & "　　" & Label2(6).Caption, "")
    strExc(1) = strExc(1) & vbCrLf & "中：" & Label2(2).Caption
    strExc(1) = strExc(1) & vbCrLf & "日：" & Label2(7).Caption
    strExc(1) = strExc(1) & vbCrLf & "開發日期：" & ChangeTStringToTDateString(m_PCU37)
    strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
       " values( '" & strUserNum & "','" & Replace(m_PCU38, ",", ";") & "',to_char(sysdate,'yyyymmdd')" & _
       ",to_char(sysdate,'hh24miss'),'潛在客戶" & Left(Label2(1).Caption, 8) & "已轉為" & m_NewNo & "','" & ChgSQL(strExc(1)) & "',null)"
    cnnConnection.Execute strSql
   'end 2018/05/16
   
   cnnConnection.CommitTrans
   FormSave = True
   
   '顯示新編號告知使用者
   MsgBox "轉入之客戶或代理人編號為 " & m_NewNo & " !", vbInformation
   
   'Added by Morgan 2019/2/26
   '複製編號至剪貼簿
   Clipboard.SetText m_NewNo
   MsgBox "編號已複製", vbInformation, MsgText(21)
   'end 2019/2/26
   
   Exit Function

ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbInformation
   Exit Function

ErrorHandler:
   cnnConnection.RollbackTrans
   Screen.MousePointer = vbDefault
   'Modify by Amy 2024/11/29 SaveXYNoSource有誤回傳其錯誤
    If stMsg = MsgText(601) Then
      stMsg = "潛在客戶轉客戶或代理人作業失敗，請洽系統管理員 !"
    End If
   MsgBox stMsg, vbCritical
   'end 2024/11/29
    
End Function

'Add by Amy 2023/05/08
Private Sub txtXYS02_GotFocus()
    InverseTextBox txtXYS02
End Sub

Private Sub txtXYS02_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtXYS02_Validate(Cancel As Boolean)
   'Modify by Amy 2024/11/29 改抓共用函數,避免有未改到
   Dim stName As String, stMsg As String
   
   If txtXYS02 = MsgText(601) Then LblSourceN.Caption = "": Exit Sub
   
   'Memo 直接按 Enter鍵存檔,部分資料未正常檢查,並調整訊息
   bCancel = False
   LblSourceN.Caption = ""
   txtXYS02 = Left(ChangeCustomerL(txtXYS02), 8) '補滿8碼
   stMsg = ChkXYSourceReason(1, Me.Name, 1, cboSource, txtXYS02, , , , , , stName)
   If stMsg <> MsgText(601) Then
      MsgBox stMsg, vbInformation
      'Memo 使用bCancel避免彈訊息後無法跳離 ex:來源選04 輸了Y編號,需刪Y編號,再重選
      bCancel = True
      txtXYS02_GotFocus
      Exit Sub
   End If
   LblSourceN.Caption = stName
   'end 2024/11/29
End Sub

Private Sub txtXYS03_GotFocus()
    InverseTextBox txtXYS03
End Sub
'end 2023/05/08
