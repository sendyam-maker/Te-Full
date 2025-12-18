VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm100101_16 
   BorderStyle     =   1  '單線固定
   Caption         =   "往來記錄資料查詢"
   ClientHeight    =   6000
   ClientLeft      =   1440
   ClientTop       =   2320
   ClientWidth     =   9310
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   9310
   Begin VB.CheckBox chkCR09 
      Caption         =   "財務處告知有產生國外交際餐費"
      Height          =   408
      Left            =   7092
      TabIndex        =   33
      Top             =   1584
      Width           =   1596
   End
   Begin VB.CommandButton cmdOpenAtt 
      Caption         =   "開啟"
      Height          =   255
      Left            =   8520
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4680
      Width           =   735
   End
   Begin VB.ListBox lstCR11 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      ItemData        =   "frm100101_16.frx":0000
      Left            =   8772
      List            =   "frm100101_16.frx":0002
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   975
      Width           =   1455
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "下一筆"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   8.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   2
      Left            =   7530
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   30
      Width           =   870
   End
   Begin VB.ListBox lstCR18 
      BeginProperty Font 
         Name            =   "細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      ItemData        =   "frm100101_16.frx":0004
      Left            =   8772
      List            =   "frm100101_16.frx":0006
      MultiSelect     =   1  '簡易多重選取
      Sorted          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1770
      Width           =   1455
   End
   Begin VB.CommandButton CmdOk1 
      Caption         =   "結束"
      Height          =   400
      Index           =   1
      Left            =   8430
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   0
      Top             =   0
      _ExtentX        =   494
      _ExtentY        =   494
      _Version        =   393216
   End
   Begin VB.TextBox txtCF 
      Height          =   300
      Index           =   6
      Left            =   4380
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5190
      Visible         =   0   'False
      Width           =   4560
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   19
      Left            =   24
      TabIndex        =   34
      Top             =   3550
      Visible         =   0   'False
      Width           =   504
      VariousPropertyBits=   671105051
      MaxLength       =   180
      Size            =   "882;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   2
      Left            =   1050
      TabIndex        =   23
      Top             =   811
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1926;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstContact 
      Height          =   600
      Left            =   1050
      TabIndex        =   32
      Top             =   1443
      Width           =   5235
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "9234;1058"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "往來日期：                           ( 西元 )"
      Height          =   180
      Index           =   13
      Left            =   90
      TabIndex        =   31
      Top             =   871
      Width           =   2685
   End
   Begin MSForms.ListBox lstUsers 
      Height          =   585
      Left            =   1050
      TabIndex        =   30
      Top             =   3277
      Width           =   1290
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "2275;1032"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   690
      Index           =   8
      Left            =   1050
      TabIndex        =   29
      Top             =   3878
      Width           =   7755
      VariousPropertyBits=   -1466941413
      ScrollBars      =   2
      Size            =   "13679;1217"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   7
      Left            =   1050
      TabIndex        =   28
      Top             =   2961
      Width           =   7755
      VariousPropertyBits=   671105051
      MaxLength       =   180
      Size            =   "13679;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   570
      Index           =   6
      Left            =   1050
      TabIndex        =   27
      Top             =   2375
      Width           =   7755
      VariousPropertyBits=   -1476378597
      MaxLength       =   200
      Size            =   "13679;1005"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   3
      Left            =   1050
      TabIndex        =   26
      Top             =   1127
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1926;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   1
      Left            =   1050
      TabIndex        =   25
      Top             =   480
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   9
      Size            =   "1926;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   5
      Left            =   1050
      TabIndex        =   24
      Top             =   2059
      Width           =   5925
      VariousPropertyBits=   671105051
      MaxLength       =   200
      Size            =   "10451;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstAtt 
      Height          =   1320
      Left            =   1050
      TabIndex        =   3
      Top             =   4620
      Width           =   7440
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "13123;2328"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.ListBox lstSort 
      Height          =   600
      Left            =   3840
      TabIndex        =   2
      Top             =   2340
      Visible         =   0   'False
      Width           =   5235
      VariousPropertyBits=   746586139
      ScrollBars      =   2
      DisplayStyle    =   2
      Size            =   "9234;1058"
      MatchEntry      =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtCR 
      Height          =   300
      Index           =   10
      Left            =   3870
      TabIndex        =   20
      Top             =   811
      Width           =   1092
      VariousPropertyBits=   671105051
      MaxLength       =   8
      Size            =   "1926;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lbl1 
      Height          =   300
      Left            =   2220
      TabIndex        =   22
      Top             =   1127
      Width           =   5010
      VariousPropertyBits=   27
      Caption         =   "lblFM2"
      Size            =   "8837;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox textCUID 
      Height          =   300
      Left            =   2190
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   480
      Width           =   6225
      VariousPropertyBits=   -2147467233
      BackColor       =   16777215
      Size            =   "10980;529"
      Caption         =   "LblFM2"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "主旨："
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   10
      Top             =   2430
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "接洽同仁："
      Height          =   180
      Index           =   11
      Left            =   90
      TabIndex        =   18
      Top             =   3330
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "被回覆紀錄編號："
      Height          =   180
      Index           =   10
      Left            =   8772
      TabIndex        =   16
      Top             =   780
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "回覆紀錄編號："
      Height          =   180
      Index           =   9
      Left            =   8772
      TabIndex        =   15
      Top             =   1620
      Width           =   1260
   End
   Begin VB.Label Label1 
      Caption         =   "往來對象："
      Height          =   180
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   1140
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "記錄編號："
      Height          =   180
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "聯絡人："
      Height          =   180
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   1455
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "往來類別："
      Height          =   180
      Index           =   3
      Left            =   90
      TabIndex        =   11
      Top             =   2085
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "場合："
      Height          =   180
      Index           =   5
      Left            =   90
      TabIndex        =   9
      Top             =   2985
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "內容："
      Height          =   180
      Index           =   6
      Left            =   90
      TabIndex        =   8
      Top             =   3900
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "附件："
      Height          =   180
      Index           =   7
      Left            =   90
      TabIndex        =   7
      Top             =   4650
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "回覆期限：                           ( 西元 )"
      Height          =   180
      Index           =   8
      Left            =   2925
      TabIndex        =   6
      Top             =   871
      Width           =   2685
   End
End
Attribute VB_Name = "frm100101_16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modified by Lydia 2022/01/10 改成Form 2.0; lstUsers、lstSort、lbl1、textCUID、txtCR(index)、lstContact、lstAtt
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'sonia 2010/8/20 日期欄已修改
'Create by Morgan 2007/12/18
Option Explicit

Public cmdState As Integer
Dim strTmp As String

Dim rsContact As ADODB.Recordset
Dim m_bReadGrid As Boolean
Dim oText As Object
Dim idx As Integer
Dim iLanguage As Integer
Dim m_bLanguage As String      '2008/12/9 ADD BY SONIA
'Modify By Sindy 2019/2/25 "CONTACTRECORD" 改為 "CONTACTFILE"
Private Const cTableName As String = "CONTACTFILE" 'Added by Lydia 2017/08/09 指定FTP資料夾名稱
Dim m_strCRexcept As String 'Added by Lydia 2025/08/08
Public m_pub_QL05 As String 'Add By Sindy 2025/8/27 只記錄於此Form


Private Sub cmdOpenAtt_Click()
'Added by Lydia 2017/08/09
Dim tmpArr As Variant, ii As Integer
Dim stFileName As String
Dim hLocalFile As Long
'end 2017/08/09
   
   'Modified by Lydia 2022/01/10 改成Form 2.0元件
   'If lstAtt.Text = "" Then
   If lstAtt.ListIndex = -1 Then
      MsgBox "請選擇欲開啟的附件！"
   Else
      'Added by Lydia 2017/08/09 判斷移檔日期
      If strSrvDate(1) >= CR_NewDate And txtCF(6).Text <> "" Then
         tmpArr = Empty
         tmpArr = Split(txtCF(6).Text, ",")
         ii = lstAtt.ListIndex
         If ii > UBound(tmpArr) Then Exit Sub
         If Trim(tmpArr(ii)) <> "" Then
            'Modified by Lydia 2022/04/18 debug
            'strExc(1) = Trim(Mid(lstAtt.Text, 1, InStrRev(lstAtt.Text, " (") - 1))
            strExc(1) = Trim(Mid(lstAtt.List(ii), 1, InStrRev(lstAtt.List(ii), " (") - 1))
            stFileName = App.path & "\$$" & strExc(1)
            If PUB_GetFtpFile(Trim(tmpArr(ii)), stFileName, cTableName) Then
                ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
            End If
         End If
      'Removed by Morgan 2024/8/2 不用的標記為註解，檢查程式碼才知時可略過
      'Else
      ''end 2017/08/09
      '    PUB_OpenFtpFile txtCR(1), lstAtt.Text, Winsock1
      'end 2024/8/2
      End If 'end 2017/08/09
   End If
End Sub

Private Sub Form_Load()
   bolToEndByNick = False
   MoveFormToCenter Me
   cmdState = -1
   textCUID.BackColor = &H8000000F
   '2008/11/10 ADD BY SONIA 回覆功能先鎖住以後再用
   Label1(8).Visible = False
   Label1(9).Visible = False
   Label1(10).Visible = False
   txtCR(10).Visible = False
   lstCR11.Visible = False
   lstCR18.Visible = False
   '2008/12/9 ADD BY SONIA
   m_bLanguage = IsUserHasRightOfLanguage
   '有值才可查潛在客戶往來記錄 Y不限語文 J限日文 E限非日文
   m_strCRexcept = Pub_GetCRExceptNo(Me.Name) 'Added by Lydia 2025/08/08
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add by Morgan 2009/5/19
   '清除暫存檔
   PUB_KillTempFile "$$*.*"
   Set frm100101_16 = Nothing
End Sub

Private Sub cmdok1_Click(Index As Integer)
   cmdState = Index
   PubShowNextData
End Sub

Public Sub PubShowNextData()
   Select Case cmdState
      Case 2
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Case 1
         fnCloseAllFrm100
   End Select
End Sub

Sub StrMenu()
   Dim strKey  As String
   strKey = Me.Tag
   pub_QL05 = m_pub_QL05 & ";記錄編號：" & Me.Tag & "(國外往來記錄資料)" 'Add By Sindy 2025/8/13
   
   strExc(0) = "select * from contactrecord,allcode where cr01='" & strKey & "'" & _
               " and ac01(+)='11' and cr05=ac02(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If pub_QL04 <> "" Then InsertQueryLog (RsTemp.RecordCount) 'Add By Sindy 2025/8/13
      'Add By Sindy 2011/01/03 檢查國內外權限
      If CheckSR12(RsTemp.Fields("cr03")) = False Then
         Screen.MousePointer = vbDefault
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
      'Added by Lydia 2025/08/08 國外往來記錄的維護及查詢限制
      If m_strCRexcept <> "" And InStr(m_strCRexcept, strKey) > 0 Then
         Screen.MousePointer = vbDefault
         MsgBox "限閱往來記錄！", vbInformation
         tmpBol = fnCancelNowFormAndShowParentForm(Me)
         Exit Sub
      End If
      'end 2025/08/08
      ShowRecord RsTemp
   Else
      If pub_QL04 <> "" Then InsertQueryLog (0) 'Add By Sindy 2025/8/13
      ShowNoData
      Screen.MousePointer = vbDefault
      tmpBol = fnCancelNowFormAndShowParentForm(Me)
      Exit Sub
   End If
End Sub

' 將資料庫中的資料更新到所有欄位中
Private Sub ShowRecord(ByRef p_Rst As ADODB.Recordset)
   Dim rsRec As ADODB.Recordset
   Dim CUID(1 To 6) As String, strName As String
   Dim AdoRs As New ADODB.Recordset 'Add By Sindy 2019/2/25
   Dim strCF02 As String 'Add By Sindy 2019/2/25
   
   ClearField
   SetCtrlReadOnly True
   Set rsRec = p_Rst.Clone
   With rsRec
      If .RecordCount > 0 Then
         '2008/12/9 ADD BY SONIA 加判斷語文權限
         'If GetCustData(Mid(.Fields("CR03"), 1, 8)) = False Then
         lbl1 = ""
         If PUB_GetCustData(.Fields("CR03"), strName) = False Then
            'Add By Sindy 2009/04/30
            'MsgBox "您沒有維護此筆潛在客戶的往來記錄權限！"
            '2009/04/30 End
            Screen.MousePointer = vbDefault
            Me.Enabled = True
            tmpBol = fnCancelNowFormAndShowParentForm(Me)
            Exit Sub
         Else
            lbl1 = strName
         End If
         '2008/12/9 END
        
         For Each oText In txtCR
            idx = oText.Index
            oText.Text = "" & .Fields("CR" & Format(idx, "0#"))
         Next
         'Add by Amy 2025/03/20 內容中若特殊符號前後框住,則依權限顯示 or 不顯示
         'ex: KA2000178/Y2005000 其客戶「#(Nabtesco)#」相關商標案件
         '[有]權限顯示:其客戶「Nabtesco」相關商標案件;[無]權限顯示:其客戶***相關商標案件
         If txtCR(8) <> MsgText(601) Then
            txtCR(8) = ChkLimitAndReplace(Me.Name, txtCR(8), txtCR(19))
         End If
         'end 2025/03/20
         
         '往來對象
         'GetCustData .Fields("CR03")
         '聯絡人
         If Not IsNull(.Fields("CR04")) Then
            setContact lstContact, .Fields("CR03"), .Fields("CR04")
         End If
         '被回覆紀錄編號
         If Not IsNull(.Fields("CR11")) Then
            SetList lstCR11, .Fields("CR11")
         End If
         '回覆紀錄編號
         If Not IsNull(.Fields("CR18")) Then
            SetList lstCR18, .Fields("CR18")
         End If
         '往來類別
'         If Not IsNull(.Fields("CR05")) Then
'            SetList lstSort, .Fields("CR05")
'         End If
         'Modify By Sindy 2019/3/8
         txtCR(5) = .Fields("CR05") & " " & .Fields("AC03")
         
         'Add by Amy 2015/06/10 +接洽同仁
         If Not IsNull(.Fields("CR19")) Then
            SetlstUsers .Fields("CR19")
         End If
         'Add By Sindy 2019/2/26
         strExc(0) = "SELECT cf02,cf06,cf07 FROM ContactFile where CF01='" & txtCR(1) & "'"
         intI = 1
         Set AdoRs = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            AdoRs.MoveFirst
            Do While Not AdoRs.EOF
               strCF02 = strCF02 & "," & AdoRs.Fields("cf02") & IIf("" & AdoRs.Fields("cf07") <> "", " (" & AdoRs.Fields("cf07") & " KB)", "")
               txtCF(6) = txtCF(6) & "," & AdoRs.Fields("cf06")
               AdoRs.MoveNext
            Loop
            strCF02 = Mid(strCF02, 2)
            txtCF(6) = Mid(txtCF(6), 2)
         Else
            strCF02 = ""
            txtCF(6) = ""
         End If
         '附件路徑
'         If Not IsNull(.Fields("CR09")) Then
'            SetList lstAtt, .Fields("CR09")
'         End If
         If Not IsNull(strCF02) Then
            SetList lstAtt, strCF02
         End If
         '2019/2/26 END
         
         CUID(1) = "" & .Fields("CR12")
         CUID(2) = "" & .Fields("CR13")
         CUID(3) = "" & .Fields("CR14")
         CUID(4) = "" & .Fields("CR15")
         CUID(5) = "" & .Fields("CR16")
         CUID(6) = "" & .Fields("CR17")
         
         'Add By Sindy 2023/8/10
         If "" & .Fields("CR09") = "Y" Then
            chkCR09.Value = 1
         Else
            chkCR09.Value = 0
         End If
         '2023/8/10 END
      End If
   End With
   UpdateCUID CUID, textCUID
   Set AdoRs = Nothing 'Add By Sindy 2019/2/25
End Sub

Private Sub ClearField()
   Dim oLabel As Object
   
   For Each oText In txtCR
      oText.Text = Empty
   Next
   lbl1 = Empty
   textCUID = ""
   lstContact.Clear
   lstSort.Clear
   lstAtt.Clear
   lstCR11.Clear
   lstCR18.Clear
   lstUsers.Clear 'Add by Amy 2015/06/10
   'Add By Sindy 2019/2/26
   For Each oText In txtCF
      oText.Text = Empty
   Next
   '2019/2/26 END
   chkCR09.Value = 0 'Add By Sindy 2023/8/10
End Sub

Private Sub SetCtrlReadOnly(ByVal bLocked As Boolean)
   For Each oText In txtCR
      oText.Locked = bLocked
   Next
   chkCR09.Enabled = Not bLocked 'Add By Sindy 2023/8/10
End Sub

'Modified by Lydia 2022/01/10 As ListBox => Object
Private Sub SetList(oList As Object, p_stList As String)
   Dim arrID
   oList.Clear
   If p_stList <> "" Then
      arrID = Split(p_stList, ",")
      For intI = UBound(arrID) To LBound(arrID) Step -1
         oList.AddItem arrID(intI), 0
      Next
   End If
End Sub

' 更新 Create 及 Update 的人
'Modified by Lydia 2022/01/10 As TextBox => Object
Private Sub UpdateCUID(ByRef p_CUID() As String, ByRef oText As Object)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If p_CUID(1) <> "" Then
      strCName = GetStaffName(p_CUID(1), True)
   End If
   If p_CUID(2) <> "" Then
      strCDate = ChangeWStringToTDateString(p_CUID(2))
   End If
   
   If p_CUID(3) <> "" Then
      strCTime = Format(p_CUID(3), "##:##")
   End If
   
   If p_CUID(4) <> "" Then
      strUName = GetStaffName(p_CUID(4), True)
   End If
   If p_CUID(5) <> "" Then
      strUDate = ChangeWStringToTDateString(p_CUID(5))
   End If
   
   If p_CUID(6) <> "" Then
      strUTime = Format(p_CUID(6), "##:##")
   End If
      
   ' 設定CUID中的文字
   oText = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
              
End Sub

'Private Function GetCustData(p_stCust As String) As Boolean
'Dim aiOrder(1 To 3) As Integer
'
'   GetCustData = False
'   '2008/12/9 modify by sonia 加國籍才能判斷語文權限
'   Select Case Left(p_stCust, 1)
'      Case "X"
'         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3 from customer where cu01='" & p_stCust & "' and cu02='0'"
'      Case "Y"
'         strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3 from fagent where fa01='" & p_stCust & "' and fa02='0'"
'      Case "R"
'         strExc(0) = "select pcu36,pcu08,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) pcu03,pcu07,PCU09 N3 from potcustomer where pcu01='" & p_stCust & "' and pcu02='0'"
'      Case Else
'         MsgBox "往來對象必須為 X、Y 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
'   End Select
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   lbl1 = ""
'   If intI = 1 Then
'      '2008/12/9 ADD BY SONIA 加語文權限
'      If m_bLanguage = "" And Left(p_stCust, 1) = "R" Then
'         MsgBox "您沒有查詢潛在客戶的往來記錄權限 !!!", vbInformation
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Exit Function
'      ElseIf m_bLanguage = "J" And Left(p_stCust, 1) = "R" And Mid(RsTemp.Fields("N3"), 1, 3) <> "011" Then
'         MsgBox "您沒有查詢英文組潛在客戶的往來記錄權限 !!!", vbInformation
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Exit Function
'      ElseIf m_bLanguage = "E" And Left(p_stCust, 1) = "R" And Mid(RsTemp.Fields("N3"), 1, 3) = "011" Then
'         MsgBox "您沒有查詢日文組潛在客戶的往來記錄權限 !!!", vbInformation
'         Screen.MousePointer = vbDefault
'         Me.Enabled = True
'         tmpBol = fnCancelNowFormAndShowParentForm(Me)
'         Exit Function
'      End If
'      '2008/12/9 END
'
'      iLanguage = Val("" & RsTemp(0))
'      Select Case iLanguage
'         Case 1 '中 -> 英 -> 日
'            aiOrder(1) = 1
'            aiOrder(2) = 2
'            aiOrder(3) = 3
'
'         Case 3 '日 -> 中 -> 英
'            aiOrder(1) = 3
'            aiOrder(2) = 1
'            aiOrder(3) = 2
'
'         Case Else '英 -> 中 -> 日
'            aiOrder(1) = 2
'            aiOrder(2) = 1
'            aiOrder(3) = 3
'      End Select
'      For intI = 1 To 3
'         If Not IsNull(RsTemp(aiOrder(intI))) Then
'            lbl1 = RsTemp(aiOrder(intI))
'            Exit For
'         End If
'      Next
'      GetCustData = True
'   End If
'End Function
'Modify By Sindy 2009/04/30
'Private Function GetCustData(p_stCust As String) As Boolean
'Dim strName As String
'
'   GetCustData = False
'
'   Select Case Left(p_stCust, 1)
'      Case "X"
'         strExc(0) = "select cu64,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,CU10 N3,CU81 from customer where cu01='" & Left(p_stCust, 8) & "' and cu02='" & Right(p_stCust, 1) & "'"
'      Case "Y"
'         strExc(0) = "select fa31,fa04,rtrim(fa05||' '||fa63||' '||fa64||' '||fa65) fa05,fa06,FA10 N3,FA46 from fagent where fa01='" & Left(p_stCust, 8) & "' and fa02='" & Right(p_stCust, 1) & "'"
'      Case "R"
'         strExc(0) = "select pcu36,pcu08,rtrim(pcu03||' '||pcu04||' '||pcu05||' '||pcu06) pcu03,pcu07,PCU09 N3,PCU41 from potcustomer where pcu01='" & Left(p_stCust, 8) & "' and pcu02='" & Right(p_stCust, 1) & "'"
'      Case Else
'         MsgBox "往來對象必須為 X、Y 或 R 開頭", vbCritical + vbOKOnly, "檢核資料"
'         Exit Function
'   End Select
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   lbl1 = ""
'   If intI = 1 Then
'      For intI = 1 To 3
'         If Not IsNull(RsTemp(intI)) Then
'            strName = RsTemp(intI)
'            Exit For
'         End If
'      Next
'
'      '依LoginUser和輸入人員之部門第一碼判斷部門權限, 相同者才可輸入查詢
'      '但M51不受限制
'      strExc(0) = "SELECT A.ST03,B.ST03 FROM STAFF A,STAFF B " & _
'                         "WHERE A.ST01 = '" & strUserNum & "' " & _
'                              "AND B.ST01 = '" & Trim(RsTemp(5)) & "' "
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If Trim(RsTemp(0)) <> "M51" And _
'            Left(Trim(RsTemp(0)), 1) <> Left(Trim(RsTemp(1)), 1) Then
'            MsgBox "您沒有維護此筆潛在客戶的往來記錄權限！"
'            Exit Function
'         End If
'      End If
'   Else
'      MsgBox "往來對象輸入錯誤！"
'      Exit Function
'   End If
'   lbl1 = strName
'
'   GetCustData = True
'End Function

'Modified by Ldyia 2022/04/18 ListBox=> Object
Private Sub setContact(oList As Object, p_stCR03 As String, p_stCR04 As String)
   Dim arrID
   oList.Clear
   Select Case iLanguage
      Case 1 '中 -> 英 -> 日
         '2008/11/18 modify by sonia 抓pcc01取cr03之前8碼
         'strExc(0) = "select pcc02 c1,nvl(pcc05,nvl(pcc03,pcc04)) c2 from potcustcont where pcc01='" & p_stCR03 & "' and instr('" & p_stCR04 & "',pcc02)>0 order by 1 desc"
         strExc(0) = "select pcc02 c1,nvl(pcc05,nvl(pcc03,pcc04)) c2 from potcustcont where pcc01='" & Mid(p_stCR03, 1, 8) & "' and instr('" & p_stCR04 & "',pcc02)>0 order by 1 desc"
      Case 3 '日 -> 英 -> 中
         '2008/11/18 modify by sonia 抓pcc01取cr03之前8碼
         'strExc(0) = "select pcc02 c1,nvl(pcc04,nvl(pcc03,pcc05)) c2 from potcustcont where pcc01='" & p_stCR03 & "' and instr('" & p_stCR04 & "',pcc02)>0 order by 1 desc"
         strExc(0) = "select pcc02 c1,nvl(pcc04,nvl(pcc03,pcc05)) c2 from potcustcont where pcc01='" & Mid(p_stCR03, 1, 8) & "' and instr('" & p_stCR04 & "',pcc02)>0 order by 1 desc"
      Case Else '英 -> 日 -> 中
         '2008/11/18 modify by sonia 抓pcc01取cr03之前8碼
         'strExc(0) = "select pcc02 c1,nvl(pcc03,nvl(pcc04,pcc05)) c2 from potcustcont where pcc01='" & p_stCR03 & "' and instr('" & p_stCR04 & "',pcc02)>0 order by 1 desc"
         strExc(0) = "select pcc02 c1,nvl(pcc03,nvl(pcc04,pcc05)) c2 from potcustcont where pcc01='" & Mid(p_stCR03, 1, 8) & "' and instr('" & p_stCR04 & "',pcc02)>0 order by 1 desc"
   End Select
      
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
         '設定聯絡人清單
         arrID = Split(p_stCR04, ",")
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("C1") = arrID(intI) Then
                  oList.AddItem "" & .Fields(1), 0
                  'oList.ItemData(0) = .Fields(0) 'Remove by Lydia 2022/01/10
                  Exit Do
               End If
               .MoveNext
            Loop
         Next
      End With
   End If
End Sub

'Add by Morgan 2009/5/20
'Modified by Lydia 2022/01/10 改成Form 2.0
'Private Sub lstAtt_DblClick()
Private Sub lstAtt_DblClick(Cancel As MSForms.ReturnBoolean)
   If cmdOpenAtt.Enabled = True Then
      cmdOpenAtt.Value = True
   End If
End Sub

'Add by Amy 2015/06/10
Private Sub SetlstUsers(p_stNums As String)
   Dim arrID
   
   lstUsers.Clear
   If p_stNums <> "" Then
      strExc(0) = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         arrID = Split(p_stNums, ",")
         With RsTemp
         '照原順序排
         For intI = UBound(arrID) To LBound(arrID) Step -1
            .MoveFirst
            Do While Not .EOF
               If .Fields("st01") = arrID(intI) Then
                  lstUsers.AddItem "" & .Fields(1), 0
                  'lstUsers.ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號 'Remove by Lydia 2022/01/10
                  .MoveLast
               End If
               .MoveNext
            Loop
         Next
         End With
      End If
   End If
End Sub


