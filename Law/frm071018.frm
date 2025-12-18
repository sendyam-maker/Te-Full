VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm071018 
   BorderStyle     =   1  '單線固定
   Caption         =   "出庭律師資料輸入"
   ClientHeight    =   5040
   ClientLeft      =   1320
   ClientTop       =   552
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   5400
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "出庭律師維護"
      Height          =   1365
      Left            =   120
      TabIndex        =   12
      Top             =   1440
      Width           =   5205
      Begin VB.CommandButton cmdCancel 
         Caption         =   "刪除"
         Height          =   315
         Left            =   4170
         TabIndex        =   4
         Top             =   548
         Width           =   795
      End
      Begin VB.CommandButton cmdInput 
         Caption         =   "增修"
         Height          =   315
         Left            =   3150
         TabIndex        =   3
         Top             =   548
         Width           =   945
      End
      Begin VB.TextBox txtCL03 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1110
         MaxLength       =   8
         TabIndex        =   2
         Top             =   960
         Width           =   1035
      End
      Begin MSForms.ComboBox cboEmp 
         Height          =   330
         Left            =   1110
         TabIndex        =   1
         Top             =   540
         Width           =   1875
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "3307;582"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "總  費  用： "
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   995
      End
      Begin MSForms.Label lblData 
         Height          =   285
         Index           =   2
         Left            =   1110
         TabIndex        =   16
         Top             =   210
         Width           =   1215
         BackColor       =   16777215
         VariousPropertyBits=   27
         Size            =   "2143;503"
         FontName        =   "新細明體"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "P.S.出庭費總合不可超過總費用"
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   2370
         TabIndex        =   15
         Top             =   240
         Width           =   2745
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "出  庭  費： "
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   14
         Top             =   1020
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "出庭律師： "
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   13
         Top             =   578
         Width           =   915
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid1 
      Height          =   2000
      Left            =   120
      TabIndex        =   11
      Top             =   2910
      Width           =   5175
      _ExtentX        =   9123
      _ExtentY        =   3514
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      FormatString    =   "V|出庭律師|出 庭 費"
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
      _Band(0).Cols   =   3
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面(&U)"
      Height          =   400
      Left            =   4200
      TabIndex        =   0
      Top             =   48
      Width           =   1155
   End
   Begin VB.Label lbeNumber 
      Height          =   270
      Left            =   1160
      TabIndex        =   19
      Top             =   150
      Width           =   2115
   End
   Begin VB.Label lbePaperNum 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   """#-##-######"""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1028
         SubFormatType   =   0
      EndProperty
      Height          =   270
      Left            =   1160
      TabIndex        =   18
      Top             =   472
      Width           =   2115
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   1
      Left            =   1160
      TabIndex        =   10
      Top             =   1116
      Width           =   2625
      VariousPropertyBits=   27
      Size            =   "4630;503"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblData 
      Height          =   285
      Index           =   0
      Left            =   1160
      TabIndex        =   9
      Top             =   794
      Width           =   4000
      VariousPropertyBits=   27
      Size            =   "7056;503"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "案件性質： "
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   8
      Top             =   1116
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "案件名稱："
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   794
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "本所案號："
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   150
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "收  文  號： "
      Height          =   255
      Index           =   1
      Left            =   180
      TabIndex        =   5
      Top             =   472
      Width           =   915
   End
End
Attribute VB_Name = "frm071018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/12/08 原名「其他出庭律師資料輸入」改名為「出庭律師資料輸入」並且增加輸入”出庭費”
'Memo by Lydia 2021/09/14 改成Form2.0 ; List1、cboEmp
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create By Sindy 2011/6/8
Option Explicit

Dim intNowRecd As Integer
Dim strTemp As Variant, i As Integer
'Modified by Lydia 2022/12/08
'Public cMainFNum As Integer
'Public UpForm As Form
Dim m_CP09 As String
Dim m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String
Dim m_PrevForm As Form '前一畫面
Dim m_NowCP14 As String  '前一畫面輸入的承辦人
Dim m_bolReset As Boolean '是否清除暫存檔=> 在確定存檔前,重複進入
Dim mESeqNo  As String  '執行序號=1, 視情況是否要改成遞增
Dim m_bolUpdate As Boolean '是否可維護資料
Dim m_bolChkFee As Boolean '是否顯示”出庭費”
Dim m_LOS02 As String '法律所案源類別
Dim m_LOS15 As String '法律所案源單號
Dim m_RowSeq As String  '暫存檔的RowSeq
Dim intLastRow As Integer '點選列
Dim unitLC03 As String '預設律師費(出庭費)
Dim rsAD As New ADODB.Recordset
Dim strA1 As String, intA As Integer
Dim colCL01 As Integer, colCL02 As Integer, colCL02n As Integer, colCL03 As Integer, colRowSeq As String
Private Const cntCL03frm = "FRM071002,FRM081002"  'Added by Lydia 2024/07/29 可維護”出庭費”的程式---分案作業
Private Const cntCPfrm = "FRM071013,FRM075012" 'Added by Lydia 2025/03/19 庭期資料/開庭通知
'Added by Lydia 2024/09/30 (113/11/01上線)
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '案源總收文號之本所案號
Dim m_LOS01fa As String  '案源之FC代理人
Dim midCP10 As String '分案作業傳入畫面的案性質
Dim bolChkReset As Boolean '是否重新預設承辦人的記錄(分案作業存檔前第2次進入)
'end 2024/09/30 (113/11/01上線)
Dim bolAddCP As Boolean 'Added by Lydia 2025/03/19 是否新增為新增庭期資料

'Added by Lydia 2022/12/08
'Modified by Lydia 2024/09/30 (113/11/01上線) +pCP10,pBolChkReset
'Modified by Lydia 2025/03/19 + pAddCP
Public Sub SetParent(ByRef fm As Form, ByVal pCP09 As String, ByVal pReset As Boolean, Optional ByVal pCP14 As String = "N", Optional ByVal pCP10 As String, Optional ByVal pBolChkReset As Boolean = False, Optional pAddCP As Boolean = False)
    Set m_PrevForm = fm
    m_CP09 = pCP09
    m_NowCP14 = pCP14
    m_bolReset = pReset
    'Added by Lydia 2024/09/30 (113/11/01上線)
    midCP10 = pCP10
    bolChkReset = pBolChkReset
    bolAddCP = pAddCP 'Added by Lydia 2025/03/19
    
End Sub

'Add By Sindy 2020/6/10
Private Sub CboEmp_Click()
   'cmdCancel.Enabled = False 'Mark by Lydia 2022/12/08
End Sub

Private Sub CboEmp_GotFocus()
   cboEmp.SelStart = 0
   cboEmp.SelLength = Len(cboEmp.Text)
End Sub

'Memo by Lydia 2021/09/14 改成Form2.0
'Private Sub CboEmp_KeyPress(KeyAscii As Integer)
Private Sub CboEmp_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboEmp_LostFocus()
Dim strText As String
   
   If cboEmp.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboEmp.Text)
      If strText <> "" Then
         cboEmp.Text = strText & " " & cboEmp.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboEmp.Text, 5))
         If strText <> "" Then
            cboEmp.Text = Left(cboEmp.Text, 5) & " " & strText
         End If
      End If
   End If
End Sub
Private Sub cboEmp_Validate(Cancel As Boolean)
   If cboEmp.Text <> "" Then
      If GetStaffName(Trim(Left(cboEmp.Text, 6)), False) = "" Then
         MsgBox "查無此員工或員工已離職！", vbInformation
         CboEmp_GotFocus
         Cancel = True
         Exit Sub
      End If
      
      'Mark by Lydia 2022/12/08
      'For i = 0 To List1.ListCount - 1
      '   If Trim(Left(List1.List(i), 6)) = Trim(Left(cboEmp.Text, 6)) Then
      '      MsgBox "此員工已輸入！", vbInformation
      '      CboEmp_GotFocus
      '      Cancel = True
      '      Exit Sub
      '   End If
      'Next
      If Val(m_RowSeq) > 0 And cboEmp.Tag <> cboEmp.Text Then
         m_RowSeq = ""
         intLastRow = 0 '修改人員，重新檢查
      End If
      If intLastRow <= 0 Then
         For i = 1 To MGrid1.Rows - 1
            If Trim(Left(cboEmp.Text, 6)) = "" & MGrid1.TextMatrix(i, colCL02) Then
                MsgBox "此員工已輸入！", vbInformation
                CboEmp_GotFocus
                Cancel = True
                Exit Sub
            End If
         Next i
      End If
      'end 2022/12/08
   End If
   If Cancel = False Then CloseIme
End Sub
'2020/6/10 END

Private Sub cmdBack_Click()
Dim strName As String
   
   'Modified by Lydia 2022/12/08
   'For i = 0 To List1.ListCount - 1
   '   strName = strName + Trim(Left(List1.List(i), 6)) + ","
   'Next
   'strPublicTemp = strName
   'Me.Hide
   'UpForm.Show
   Me.Hide
   m_PrevForm.Show
   'end 2022/12/08
   Unload Me
End Sub

Private Sub cmdCancel_Click()
   'Modified by Lydia 2022/12/08
   'If List1.ListCount > 0 Then
   '   List1.RemoveItem intNowRecd
   '   'Modify By Sindy 2020/6/10
   '   'txtReceiver = ""
   '   cboEmp.Text = ""
   '   '2020/6/10 END
   'End If
    'cmdCancel.Enabled = False
   If m_RowSeq <> "" Then
       strSql = "delete from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and rowseq=" & CNULL(m_RowSeq)
       cnnConnection.Execute strSql, intI
   End If
   Call ReadTemp(False)
   'end 2022/12/08
End Sub

Private Sub cmdInput_Click()
Dim Cancel As Boolean

   'cmdCancel.Enabled = False 'Mark by Lydia 2022/12/08
   'Modify By Sindy 2020/6/10 txtReceiver => cboEmp
   'If txtReceiver <> "" Then
   If Trim(cboEmp.Text) <> "" Then
      Cancel = False
      'Modify By Sindy 2020/6/10 txtReceiver => cboEmp
      Call cboEmp_Validate(Cancel)
      If Cancel = False Then
         'Modified by Sindy 2020/6/10
         'List1.AddItem Trim(Left(txtReceiver, 6)) & " " & GetStaffName(Trim(Left(txtReceiver, 6)), False)
         'txtReceiver = ""
         'txtReceiver.SetFocus
         'Modified by Lydia 2022/12/08 先存暫存檔
         'List1.AddItem Trim(Left(cboEmp.Text, 6)) & " " & GetStaffName(Trim(Left(cboEmp.Text, 6)), False)
         If Val(txtCL03) = 0 And txtCL03.Visible = True Then
            'Added by Lydia 2024/09/30 (113/11/01上線) 出庭費可以輸入0的狀況:
            '1-增加特定案件性質可輸出庭費，但也可以輸0表示有輸過。
            '2-案源為商標且有FC代理人之法務案34行政訴訟程序若已輸入0則不必再提醒。
            '但其他情形則一定要輸入金額。
            If (m_LOS01cp01 <> "TT" And InStr(m_LOS01cp01, "T") > 0 And m_LOS01fa <> "" And midCP10 = "34") Or (InStr(";" & Pub_GetSpecMan("出庭費特殊性質") & ";", ";" & midCP10 & ";") > 0) Then
            Else
            'end 2024/09/30 (113/11/01上線)
               If Val(lblData(2)) > 0 Then
                  MsgBox "出庭費不可為0　！！", vbCritical
                  GoTo EXITSUB
               End If
            End If 'Added by Lydia 2024/09/30 (113/11/01上線)
         End If
         
         strExc(0) = "select max(rowseq) as mno,sum(r003) as tfee from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If Val("" & RsTemp.Fields("tfee")) + Val(txtCL03) - Val(txtCL03.Tag) > Val(lblData(2)) Then
                MsgBox "出庭費總合不可超過總費用！", vbCritical, "檢核資料"
                GoTo EXITSUB
            End If
            If Val(m_RowSeq) > 0 Then
                strSql = "Update rdatafactory set r005=" & CNULL(Trim(Left(cboEmp.Text, 6))) & ", r003=" & CNULL(txtCL03) & _
                           " where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and rowseq=" & CNULL(m_RowSeq)
            Else
                'Added by Lydia 2024/02/17 無資料=無序號
                If Val(mESeqNo) = 0 Then
                   mESeqNo = "1"
                End If
                'end 2024/02/17
                strSql = "insert into  rdatafactory (FORMNAME,ID,SEQNO,ROWSEQ,R001,R002,R003,R004,R005,R006) values ('" & Me.Name & "'," & _
                           " '" & strUserNum & "', '" & mESeqNo & "', '" & Val("" & RsTemp.Fields("mno")) + 1 & "', '','" & Trim(cboEmp.Text) & "'," & _
                           " '" & Val(txtCL03) & "', '" & m_CP09 & "', '" & Trim(Left(cboEmp.Text, 6)) & "', '2') "
            End If
            cnnConnection.Execute strSql, intI
         End If
         Call ReadTemp(False)
         'end 2022/12/08
         cboEmp.Text = ""
         cboEmp.SetFocus
         '2020/6/10 END
      End If
   End If
   '2020/6/10 END
   
'Added by Lydia 2022/12/08
   Exit Sub
EXITSUB:
   txtCL03.SetFocus
   Call txtCL03_GotFocus
End Sub

Private Sub Form_Activate()
'Modified by Sindy 2020/6/10
'   If txtReceiver.Visible = True Then
'      txtReceiver.SetFocus
'   End If
   If cboEmp.Visible = True Then
      cboEmp.SetFocus
   End If
'2020/6/10 END
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me
   'cmdCancel.Enabled = False 'Mark by Lydia 2022/12/08
   
   'Add By Sindy 2020/6/10 出庭律師增加下拉功能，預設L01部門人員，依所別＋員工代號排序
   '出庭律師
   cboEmp.Clear
   cboEmp.AddItem ""
   strSql = "SELECT st01,st02 FROM staff WHERE st04='1' and st03='L01' order by st06 asc,st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            cboEmp.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   '2020/6/10 END
    
   'Modified by Lydia 2022/12/08
   'ReadTemp
   '部門為LXX及ST05 in ('01','08','09','00')的人才可以出現”出庭費”
   'Modified by Lydia 2023/01/06 設定可看開庭費及輸入開庭費之法律所同仁為系統特殊設定「出庭費維護」，其他法律所同仁請關閉權限
   'If Left(Pub_StrUserSt03, 1) = "L" Or InStr("01,08,09,00", Pub_strUserST05) > 0 Then
   If (Left(Pub_StrUserSt03, 1) = "L" And InStr(Pub_GetSpecMan("出庭費維護"), strUserNum) > 0) Or InStr("01,08,09,00", Pub_strUserST05) > 0 Then
       m_bolChkFee = True
   End If
   
   'Modified by Lydia 2024/07/29 改成常數cntCL03frm
   'If m_bolChkFee = True And InStr("FRM071002,FRM081002,FRM071013", UCase(TypeName(m_PrevForm))) > 0 Then
   'Modified by Lydia 2025/03/19 改成常數 cntCPfrm
   'If m_bolChkFee = True And InStr(cntCL03frm & ",FRM071013,FRM075012", UCase(TypeName(m_PrevForm))) > 0 Then
   If InStr(cntCL03frm & "," & cntCPfrm, UCase(TypeName(m_PrevForm))) > 0 Then
      m_bolUpdate = True
   End If
   If m_bolUpdate = False Then
       Frame1.Visible = False
       MGrid1.Top = Frame1.Top
       MGrid1.Height = 3375
   Else
       '分案作業呼叫才可維護”出庭費”
       'Modified by Lydia 2024/07/29 改成常數cntCL03frm,排除非分案作業
       'If InStr("FRM071013", UCase(TypeName(m_PrevForm))) > 0 Then
       If InStr(cntCL03frm, UCase(TypeName(m_PrevForm))) = 0 Then
            Label1(2).Visible = False: lblData(2).Visible = False: Label3.Visible = False
            Label1(4).Visible = False: txtCL03.Visible = False
       End If
       Frame1.Visible = True
   End If
   strSql = "select cp01,cp02,cp03,cp04,cp09,cp10,nvl(lc05,nvl(lc06,lc07)) casename,decode(lc15,'000',cpm03,cpm04) cpm0304,los02,los15,lc15 " & _
               "from caseprogress,lawcase,casepropertymap,lawofficesource " & _
               "where cp09='" & m_CP09 & "' and cp01=lc01(+) and cp02=lc02(+) and cp03=lc03(+) and cp04=lc04(+) and cp01=cpm01(+) and cp10=cpm02(+) and cp162=los15(+) "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      lbeNumber = GiveSymbol(RsTemp.Fields("cp01"), RsTemp.Fields("cp02"), RsTemp.Fields("cp03"), RsTemp.Fields("cp04"))
      m_CP01 = "" & RsTemp.Fields("CP01")
      m_CP02 = "" & RsTemp.Fields("CP02")
      m_CP03 = "" & RsTemp.Fields("CP03")
      m_CP04 = "" & RsTemp.Fields("CP04")
      lbePaperNum = m_CP09
      lblData(0) = "" & RsTemp.Fields("casename")
      lblData(1) = "" & RsTemp.Fields("cpm0304")
      m_LOS02 = "" & RsTemp.Fields("los02")
      m_LOS15 = "" & RsTemp.Fields("los15")
      'Added by Lydia 2024/08/26 顯示出庭費,只要排除來自庭期資料維護
      'Mark by Lydia 2025/03/20 不用特別隱藏---秀玲
      'If m_CP01 <> "LA" And "" & RsTemp.Fields("cp10") = "9001" Then
      '   m_bolChkFee = False
      'End If
      'end 2024/08/26
      'end 2025/03/20
      'Added by Lydia 2024/09/30 (113/11/01上線) 分案作業傳入畫面的案性質
      If "" & RsTemp.Fields("cp10") <> midCP10 And midCP10 <> "" Then
         Call ClsPDGetCaseProperty(m_CP01, midCP10, strExc(1), IIf("" & RsTemp.Fields("lc15") <> "000", True, False))
         lblData(1) = strExc(1)
      Else
         midCP10 = "" & RsTemp.Fields("cp10")
      End If
      'end 2024/09/30 (113/11/01上線)
   End If
   'Added by Lydia 2024/09/30 (113/11/01上線) 案源資料
   If m_LOS15 <> "" Then
      strSql = "select los01,c1.cp01,c1.cp02,c1.cp03,c1.cp04,nvl(tm44,pa75) as fano " & _
               "from lawofficesource,caseprogress c1, trademark ,patent where los15='" & m_LOS15 & "' and los01=c1.cp09(+) " & _
               "and c1.cp01=tm01(+) and c1.cp02=tm02(+) and c1.cp03=tm03(+) and c1.cp04=tm04(+) and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If "" & RsTemp.Fields("cp01") <> "TT" Then
            m_LOS01cp01 = "" & RsTemp.Fields("cp01")
            m_LOS01cp02 = "" & RsTemp.Fields("cp02")
            m_LOS01cp03 = "" & RsTemp.Fields("cp03")
            m_LOS01cp04 = "" & RsTemp.Fields("cp04")
            m_LOS01fa = "" & RsTemp.Fields("fano")
         End If
      End If
   End If
   'end 2024/09/30 (113/11/01上線)
   
   lblData(2) = "0"
   'Modified by Lydia 2024/07/29 改成常數cntCL03frm,限制分案作業
   'If m_bolUpdate = True And InStr("FRM071013", UCase(TypeName(m_PrevForm))) = 0 Then
   'Modified by Lydia 2025/03/19
   'If m_bolUpdate = True And InStr(cntCL03frm, UCase(TypeName(m_PrevForm))) > 0 Then
   If m_bolChkFee = True And InStr(cntCL03frm, UCase(TypeName(m_PrevForm))) > 0 Then
      If m_CP01 = "FCL" And m_LOS02 > "B" Then '外法B類案源的費用在PT案
         strSql = "select (nvl(cp16,0)-sum(nvl(a1u07,0))-sum(nvl(a1u09,0)))-(nvl(cp17,0)-sum(nvl(a1u09,0))) amt1 " & _
                      "from caseprogress,acc1u0 where cp09=(select los01 from caseprogress,lawofficesource where cp09='" & m_CP09 & "' and cp162=los15(+) and los15 is not null) " & _
                      "and cp09<'C' and nvl(cp16,0)>0 and cp09=a1u03(+) group by cp16,cp17 "
      Else
         strSql = "select (nvl(cp16,0)-sum(nvl(a1u07,0))-sum(nvl(a1u09,0)))-(nvl(cp17,0)-sum(nvl(a1u09,0))) amt1 " & _
                      "from caseprogress,acc1u0 where cp09='" & m_CP09 & "'  and cp09<'C' and nvl(cp16,0)>0 and cp09=a1u03(+) group by cp16,cp17 "
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         lblData(2) = Val("" & RsTemp.Fields("amt1"))
      End If
   End If
   Call ReadTemp(m_bolReset)
   'end 2022/12/08
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Modified by Lydia 2024/07/29 改成常數cntCL03frm
   'If InStr("FRM071002,FRM081002,FRM071013", UCase(TypeName(m_PrevForm))) > 0 Then
   'Modified by Lydia 2024/07/29 改成常數cntCPfrm
   If InStr(cntCL03frm & "," & cntCPfrm, UCase(TypeName(m_PrevForm))) > 0 Then
      strPublicTemp = ""
      For i = 1 To MGrid1.Rows - 1
         If MGrid1.TextMatrix(i, colCL02) <> "" Then 'Added by Lydia 2024/09/30 (113/11/01上線)
             strPublicTemp = strPublicTemp & MGrid1.TextMatrix(i, colCL02) & ","
         End If
      Next i
      'Added by Lydia 2024/09/30 (113/11/01上線)
      If strPublicTemp = "" Then
         strSql = "delete from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and rowseq=" & CNULL(m_RowSeq)
         cnnConnection.Execute strSql, intI
      Else
      'end 2024/09/30 (113/11/01上線)
         m_PrevForm.Tag = Me.Name & "|" & mESeqNo
      End If
   End If
   Set frm071018 = Nothing
End Sub

'Mark by Lydia 2022/12/08
'Private Sub List1_Click()
'   cmdCancel.Enabled = True
'   intNowRecd = List1.ListIndex
'   'Modified by Sindy 2020/6/10
'   'txtReceiver = List1.Text
'   'If txtReceiver <> "" And txtReceiver.Visible = True Then cmdCancel.SetFocus
'   cboEmp = List1.Text
'   If cboEmp <> "" And cboEmp.Visible = True Then cmdCancel.SetFocus
'   '2020/6/10 END
'End Sub
'end 2022/12/08

'Modified by Lydia 2022/12/08 +  bolReset
Private Sub ReadTemp(ByVal bolReset As Boolean)
   'Modified by Lydia 2022/12/08
   'strTemp = Split(strPublicTemp, ",")
   'For i = 0 To UBound(strTemp) - 1
   '   List1.AddItem strTemp(i) & " " & GetStaffName(strTemp(i), True)
   'Next i
   Call SetGrid(True)

   If bolReset = True Then
JumpToReset:  'Added by Lydia 2024/09/30 (113/11/01上線)
      cnnConnection.Execute "delete from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' "
      'Memo by Lydia 2024/08/26 只有在輸入來函-庭期資料維護,沒有收文號
      'Modified by Lydia 2025/03/19 改用變數控制
      'If m_bolUpdate = True And InStr("FRM071013", UCase(TypeName(m_PrevForm))) > 0 Then
      If bolAddCP = True Then
           strA1 = "SELECT '' AS V,ST01||' '||ST02 AS CL02N, 0 AS CL03, 'AAAA' AS CL01,'" & m_NowCP14 & "' AS CL02,'0' ORD1 FROM STAFF WHERE ST01=" & CNULL(m_NowCP14)
           m_CP09 = "AAAA"
           mESeqNo = "1"
      Else
           'Modified by Lydia 2023/08/14 debug:有設承辦人,但又沒有出庭律師記錄 + AND NVL(CP14,'N') <> 'N'
           'Modified by Lydia 2024/09/30 (113/11/01上線) 有CaseLawer就不預設承辦人的記錄
           'strA1 = "SELECT '' AS V,NVL(CP14,'" & m_NowCP14 & "')||' '||ST02 AS CL02N, 0 AS CL03, CP09 AS CL01,NVL(CP14,'" & m_NowCP14 & "') AS CL02,DECODE(CP14,NULL,'0','1') ORD1 " & _
                       "FROM CASEPROGRESS,STAFF WHERE CP09='" & m_CP09 & "' AND NVL(CP14,'" & m_NowCP14 & "')=ST01(+) " & _
                       "AND NVL(CP14,'N') <> 'N' AND NVL(CP14,'" & m_NowCP14 & "') NOT IN (SELECT CL02 FROM CASELAWER WHERE CL01='" & m_CP09 & "') " & _
                       "UNION SELECT '' AS V,ST01||' '||ST02 AS CL02N, NVL(CL03,0) AS CL03,CL01,CL02,'2' ORD2 FROM CASELAWER,STAFF WHERE CL01='" & m_CP09 & "' AND CL02=ST01(+) "
           strA1 = "SELECT '' AS V,NVL(CP14,'" & m_NowCP14 & "')||' '||ST02 AS CL02N, 0 AS CL03, CP09 AS CL01,NVL(CP14,'" & m_NowCP14 & "') AS CL02,DECODE(CP14,NULL,'0','1') ORD1 " & _
                       "FROM CASEPROGRESS,STAFF WHERE CP09='" & m_CP09 & "' AND NVL(CP14,'" & m_NowCP14 & "')=ST01(+) " & _
                       "AND NVL(CP14,'N') <> 'N' AND CP09 NOT IN (SELECT CL01 FROM CASELAWER WHERE CL01='" & m_CP09 & "') " & _
                       "UNION SELECT '' AS V,ST01||' '||ST02 AS CL02N, NVL(CL03,0) AS CL03,CL01,CL02,'2' ORD2 FROM CASELAWER,STAFF WHERE CL01='" & m_CP09 & "' AND CL02=ST01(+) "
      End If
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strA1)
      If intA = 1 Then
          Set RsTemp = PUB_CreateRecordset(rsAD, , , , Me.Name, mESeqNo)
          '比照PUB_UpdateTTFee
          If m_LOS02 = "A4" Or m_LOS02 = "B1" Then
                 unitLC03 = 非B2律師費
          ElseIf m_LOS02 = "B2" Then
              If PUB_IsB2NeedCourt(m_LOS15) = True Then
                  unitLC03 = B2律師費
              End If
          'Added by Lydia 2023/03/15 現行未預設出庭費者，一律預設15000。
          Else
              unitLC03 = "15000"
          'end 2023/03/15
          End If
          '分案作業呼叫才可維護”出庭費”
          'Modified by Lydia 2024/05/10 查詢模式：對於沒有在CaseLawer輸入出庭費，不預設金額
          'If InStr("FRM071013", UCase(TypeName(m_PrevForm))) > 0 Then
          'Modified by Lydia 2024/07/29 改成常數cntCL03frm,排除非分案作業
          'If InStr("FRM071002,FRM081002", UCase(TypeName(m_PrevForm))) = 0 Then
          If InStr(cntCL03frm, UCase(TypeName(m_PrevForm))) = 0 Then
             unitLC03 = ""
          End If
          If Val(unitLC03) > 0 Then
              'Modified by Lydia 2024/09/12 出庭費可以輸入0，加判斷+ and R006<>'2'
              strSql = "Update Rdatafactory set R003=" & unitLC03 & " where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and nvl(R003,'0') = '0' and R006<>'2' "
              cnnConnection.Execute strSql
          'Added by Lydia 2024/05/10
          'Modified by Lydia 2024/07/29
          'Else
          ElseIf InStr(cntCL03frm, UCase(TypeName(m_PrevForm))) > 0 Then
              strA1 = "SELECT '' AS V,NVL(CP14,'" & m_NowCP14 & "')||' '||ST02 AS CL02N, 0 AS CL03, CP09 AS CL01,NVL(CP14,'" & m_NowCP14 & "') AS CL02,DECODE(CP14,NULL,'0','1') ORD1 " & _
                       "FROM CASEPROGRESS,STAFF WHERE CP09='" & m_CP09 & "' AND NVL(CP14,'" & m_NowCP14 & "')=ST01(+) " & _
                       "AND NVL(CP14,'N') <> 'N' AND NVL(CP14,'" & m_NowCP14 & "') NOT IN (SELECT CL02 FROM CASELAWER WHERE CL01='" & m_CP09 & "') "
              intA = 1
              Set rsAD = ClsLawReadRstMsg(intA, strA1)
              If intA = 1 Then
                  MsgBox "承辦人:" & rsAD.Fields("cl02n") & vbCrLf & "尚未輸入出庭費！", vbInformation
              End If
          'end 2024/05/10
          End If
      End If
   Else
      mESeqNo = "1"
      'Added by Lydia 2024/09/30 (113/11/01上線) 從分案作業第2次以後進入，沒有輸入CaseLawer都要重新預設承辦人+出庭費
      If bolChkReset = True Then
         bolChkReset = False
         strA1 = "select R002,R003 from rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' "
         intA = 1
         Set rsAD = ClsLawReadRstMsg(intA, strA1)
         If intA = 0 Then
            GoTo JumpToReset
         End If
      End If
      'end 2024/09/30 (113/11/01上線)
   End If
   
   Call ClearData
   
   'Added by Lydia 2024/09/12 承辦人若無出庭費資料，不要預設承辦人在下方GRID中，這樣會以為已經有輸入，可預設在輸入的欄位上
   strA1 = "SELECT R001 AS V, R002 AS 出庭律師, R003 AS 出庭費, R004 AS CL01, R005 AS CL02, R006 AS ORD1,ROWSEQ " & _
           "FROM RDATAFACTORY where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and r006='1' and r005='" & m_NowCP14 & "' "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      cboEmp.Text = "" & rsAD.Fields("出庭律師")
      txtCL03 = "" & rsAD.Fields("出庭費")
      strSql = "Delete from Rdatafactory where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' and r006='1' and r005='" & m_NowCP14 & "' "
      cnnConnection.Execute strSql
   End If
   'end 2024/09/12
   
   intLastRow = 0
   strA1 = "SELECT R001 AS V, R002 AS 出庭律師, R003 AS 出庭費, R004 AS CL01, R005 AS CL02, R006 AS ORD1,ROWSEQ " & _
               "FROM RDATAFACTORY where id = '" & strUserNum & "' and formname='" & Me.Name & "' and seqno='" & mESeqNo & "' " & _
               "order by r004, r005 "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
        Set MGrid1.Recordset = rsAD
        Call SetGrid(False)
   End If
   'end 2022/12/08
End Sub

'Added by Lydia 2022/12/08
Private Sub SetGrid(Optional ByVal pReset As Boolean = False)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer, intX As Integer
   
   arrGridHeadText = Array("V", "出庭律師", "出  庭  費", "CL01", "CL02", "ORD1", "ROWSEQ")
   'Modified by Lydia 2024/07/29 改成常數---區別:開庭通知
   'If m_bolChkFee = False Or InStr("FRM071013", UCase(TypeName(m_PrevForm))) > 0 Then
   'Modified by Lydia 2024/08/26 debug: 只要排除來自庭期資料維護
   'If m_bolChkFee = False Or InStr(cntCL03frm, UCase(TypeName(m_PrevForm))) = 0 Then
   'Modified by Lydia 2025/03/19 改成常數cntCPfrm
   If m_bolChkFee = False Or InStr(cntCPfrm, UCase(TypeName(m_PrevForm))) > 0 Then
       arrGridHeadWidth = Array(200, 1200, 0, 0, 0, 0, 0)
   Else
       arrGridHeadWidth = Array(200, 1200, 920, 0, 0, 0, 0)
   End If
   
   MGrid1.Visible = False
   MGrid1.Cols = UBound(arrGridHeadText) + 1
   If pReset = True Then
         MGrid1.Clear
         MGrid1.Rows = 2
   End If
       
   For iRow = 0 To MGrid1.Cols - 1
       MGrid1.row = 0
       MGrid1.col = iRow
       MGrid1.Text = arrGridHeadText(iRow)
       MGrid1.CellAlignment = flexAlignCenterCenter
       MGrid1.ColWidth(iRow) = arrGridHeadWidth(iRow)
   Next

   'Mgrid1的特定欄位之位置
   If colCL01 = 0 Then
       colCL01 = PUB_MGridGetId("CL01", MGrid1)
       colCL02 = PUB_MGridGetId("CL02", MGrid1)
       colCL02n = PUB_MGridGetId("出庭律師", MGrid1)
       colCL03 = PUB_MGridGetId("出  庭  費", MGrid1)
       colRowSeq = PUB_MGridGetId("ROWSEQ", MGrid1)
   End If
   For intI = 1 To MGrid1.Rows - 1
        MGrid1.row = intI
        MGrid1.col = colCL03
        MGrid1.CellAlignment = flexAlignCenterCenter
   Next intI
   MGrid1.Visible = True
End Sub

Private Sub MGrid1_Click()

   'intLastRow = MGrid1.MouseRow '方便中斷的偵測,但是會影響Grid單選控制
   GridClick MGrid1, intLastRow, 0, 0
   
   If intLastRow > 0 Then
      ReadGrid
   End If
End Sub

Private Sub ClearData()
   cboEmp.Text = "": cboEmp.Tag = ""
   txtCL03.Text = "": txtCL03.Tag = ""
   m_RowSeq = ""
End Sub

Private Sub ReadGrid()
   Call ClearData
   If intLastRow > 0 Then
       m_RowSeq = "" & MGrid1.TextMatrix(intLastRow, colRowSeq)
       cboEmp.Text = "" & MGrid1.TextMatrix(intLastRow, colCL02n)
       txtCL03.Text = "" & MGrid1.TextMatrix(intLastRow, colCL03)
       cboEmp.Tag = cboEmp.Text
       txtCL03.Tag = txtCL03.Text
   End If
End Sub

Private Sub txtCL03_GotFocus()
   TextInverse txtCL03
End Sub

Private Sub txtCL03_Validate(Cancel As Boolean)
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtCL03) = False Then
      If IsNumeric(txtCL03) = False Then
         Cancel = True
         strTit = "檢核資料"
         strMsg = "費用只可輸入數值資料"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtCL03_GotFocus
      End If
   End If
End Sub

